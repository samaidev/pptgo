package pptgo

// 本文件实现演示文稿的读取：现代 .pptx 支持读取与编辑，
// 旧版 .ppt 仅支持读取。

import (
	"fmt"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/samaidev/officecore"
	"github.com/samaidev/officecore/cfb"
	"github.com/samaidev/officecore/ooxml"
)

// === 现代格式：.pptx ===

// Deck 表示一个已打开的 .pptx 文件。
//
// 与 Presentation 不同，Deck 保留了原始包的全部内容。修改文本后保存时，
// 只有被改动的幻灯片会被重写，母版、版式、主题、图片、动画等一律保持原样。
type Deck struct {
	pkg *ooxml.Package

	// slideParts 是各幻灯片部件的路径，按幻灯片编号排序。
	slideParts []string

	// bodies 保存各幻灯片 XML 的当前内容，与 slideParts 一一对应。
	bodies []string
}

// slideNumRe 用于从部件路径中提取幻灯片编号。
var slideNumRe = regexp.MustCompile(`slide(\d+)\.xml$`)

// OpenPPTX 以保真模式打开一个 .pptx 文件。
func OpenPPTX(path string) (*Deck, error) {
	pkg, err := ooxml.Open(path)
	if err != nil {
		return nil, fmt.Errorf("pptgo: %w", err)
	}
	return deckFromPackage(pkg)
}

// OpenPPTXBytes 从内存中的字节流打开 .pptx。
func OpenPPTXBytes(data []byte) (*Deck, error) {
	pkg, err := ooxml.OpenBytes(data)
	if err != nil {
		return nil, fmt.Errorf("pptgo: %w", err)
	}
	return deckFromPackage(pkg)
}

// deckFromPackage 由已加载的包构造 Deck。
func deckFromPackage(pkg *ooxml.Package) (*Deck, error) {
	if !pkg.Has("ppt/presentation.xml") {
		return nil, fmt.Errorf("pptgo: 缺少 ppt/presentation.xml，这不是一个演示文稿")
	}

	var parts []string
	for _, name := range pkg.NamesWithPrefix("ppt/slides/") {
		// 排除 _rels 目录下的关系文件。
		if strings.Contains(name, "/_rels/") || !strings.HasSuffix(name, ".xml") {
			continue
		}
		parts = append(parts, name)
	}

	// 按幻灯片编号排序。字典序会把 slide10 排在 slide2 之前，
	// 因此必须按数值比较。
	sort.Slice(parts, func(i, j int) bool {
		return slideNumOf(parts[i]) < slideNumOf(parts[j])
	})

	d := &Deck{pkg: pkg, slideParts: parts}
	for _, p := range parts {
		d.bodies = append(d.bodies, string(pkg.Part(p)))
	}

	return d, nil
}

// slideNumOf 从部件路径中提取幻灯片编号，无法识别时返回 0。
func slideNumOf(part string) int {
	m := slideNumRe.FindStringSubmatch(part)
	if m == nil {
		return 0
	}
	n, err := strconv.Atoi(m[1])
	if err != nil {
		return 0
	}
	return n
}

// SlideCount 返回幻灯片数量。
func (d *Deck) SlideCount() int { return len(d.slideParts) }

// SlideText 返回第 index 张幻灯片的文本，下标从 0 开始。
//
// 同一段落中的多个 run 会被合并，因此返回的是完整语句而非零散片段。
func (d *Deck) SlideText(index int) (string, error) {
	if index < 0 || index >= len(d.bodies) {
		return "", fmt.Errorf("pptgo: 幻灯片下标 %d 越界（共 %d 张）", index, len(d.bodies))
	}
	blocks := ooxml.ExtractTextBlocks([]byte(d.bodies[index]), "p", "t")
	return strings.Join(blocks, "\n"), nil
}

// Text 返回整个演示文稿的文本，幻灯片之间以空行分隔。
func (d *Deck) Text() (string, error) {
	var sb strings.Builder
	for i := range d.bodies {
		if i > 0 {
			sb.WriteString("\n")
		}
		text, err := d.SlideText(i)
		if err != nil {
			return "", err
		}
		if len(d.bodies) > 1 {
			sb.WriteString(fmt.Sprintf("--- 第 %d 页 ---\n", i+1))
		}
		sb.WriteString(text)
		sb.WriteString("\n")
	}
	return sb.String(), nil
}

// atRe 匹配 DrawingML 的 <a:t> 文本元素。
var atRe = regexp.MustCompile(`(?s)<a:t>(.*?)</a:t>`)

// ReplaceText 替换所有幻灯片中的文本，返回替换次数。
//
// 替换按 run 粒度进行，因此字体、颜色、大小等格式会被保留。
// 若目标文本恰好跨越 run 边界，则不会被匹配到。
func (d *Deck) ReplaceText(old, new string) int {
	if old == "" {
		return 0
	}

	count := 0
	escapedOld := ooxml.EscapeXMLText(old)
	escapedNew := ooxml.EscapeXMLText(new)

	for i, body := range d.bodies {
		d.bodies[i] = atRe.ReplaceAllStringFunc(body, func(m string) string {
			sub := atRe.FindStringSubmatch(m)
			if sub == nil {
				return m
			}
			n := strings.Count(sub[1], escapedOld)
			if n == 0 {
				return m
			}
			count += n
			return "<a:t>" + strings.ReplaceAll(sub[1], escapedOld, escapedNew) + "</a:t>"
		})
	}

	return count
}

// SaveAs 将演示文稿写入 path。
// 仅被修改过的幻灯片会被重写，其余部件按原始字节复制。
func (d *Deck) SaveAs(path string) error {
	d.flush()
	if err := d.pkg.Save(path); err != nil {
		return fmt.Errorf("pptgo: %w", err)
	}
	return nil
}

// SaveBytes 返回演示文稿序列化后的字节流。
func (d *Deck) SaveBytes() ([]byte, error) {
	d.flush()
	data, err := d.pkg.WriteBytes()
	if err != nil {
		return nil, fmt.Errorf("pptgo: %w", err)
	}
	return data, nil
}

// flush 将内存中的幻灯片内容写回包。
func (d *Deck) flush() {
	for i, part := range d.slideParts {
		d.pkg.SetPart(part, []byte(d.bodies[i]))
	}
}

// === 旧版格式：.ppt ===

// PowerPoint Document 流中的记录类型。
const (
	recTextCharsAtom = 0x0FA0 // UTF-16LE 文本
	recTextBytesAtom = 0x0FA8 // 单字节压缩文本
	recCString       = 0x0FBA // UTF-16LE 字符串（大纲、标题）
)

// LegacyDeck 表示一个只读的 .ppt 演示文稿。
type LegacyDeck struct {
	// texts 是从演示文稿中提取的文本片段，按出现顺序排列。
	texts []string
}

// OpenLegacy 读取一个 .ppt 文件。
//
// 返回的演示文稿是只读的。.ppt 的记录结构不携带可靠的幻灯片归属信息，
// 因此仅能按文档顺序返回文本片段，无法准确还原分页。
func OpenLegacy(path string) (*LegacyDeck, error) {
	c, err := cfb.Open(path)
	if err != nil {
		return nil, fmt.Errorf("pptgo: %w", err)
	}
	return parseLegacyDeck(c)
}

// OpenLegacyBytes 从内存中的字节流读取 .ppt。
func OpenLegacyBytes(data []byte) (*LegacyDeck, error) {
	c, err := cfb.OpenReader(newByteReaderAt(data))
	if err != nil {
		return nil, fmt.Errorf("pptgo: %w", err)
	}
	return parseLegacyDeck(c)
}

// parseLegacyDeck 从复合文档中提取文本。
func parseLegacyDeck(c *cfb.Container) (*LegacyDeck, error) {
	stream, err := c.MustStream("PowerPoint Document")
	if err != nil {
		return nil, fmt.Errorf("pptgo: %w", err)
	}

	return &LegacyDeck{texts: extractPPTText(stream)}, nil
}

// extractPPTText 遍历记录树并收集文本原子。
//
// 记录头为 8 字节：2 字节版本与实例号、2 字节类型、4 字节长度。
// 版本字段的低 4 位为 0xF 表示这是一个容器，其负载由子记录构成，
// 需要递归进入；否则为原子记录，直接读取其负载。
func extractPPTText(stream []byte) []string {
	var out []string
	walkPPTRecords(stream, &out, 0)
	return out
}

// walkPPTRecords 递归遍历记录树。depth 用于限制递归深度，
// 防止构造异常的文件导致栈溢出。
func walkPPTRecords(data []byte, out *[]string, depth int) {
	if depth > 32 {
		return
	}

	for i := 0; i+8 <= len(data); {
		verInst := uint16(data[i]) | uint16(data[i+1])<<8
		typ := uint16(data[i+2]) | uint16(data[i+3])<<8
		length := int(uint32(data[i+4]) | uint32(data[i+5])<<8 |
			uint32(data[i+6])<<16 | uint32(data[i+7])<<24)

		i += 8
		if length < 0 || i+length > len(data) {
			// 记录被截断，停止解析。
			return
		}
		payload := data[i : i+length]
		i += length

		if verInst&0x0F == 0x0F {
			// 容器记录：递归处理其子记录。
			walkPPTRecords(payload, out, depth+1)
			continue
		}

		switch typ {
		case recTextBytesAtom:
			if s := cleanPPTText(decodeTextBytes(payload)); s != "" {
				*out = append(*out, s)
			}

		case recTextCharsAtom, recCString:
			// UTF-16LE。
			if len(payload)%2 != 0 {
				payload = payload[:len(payload)-1]
			}
			u16 := make([]uint16, len(payload)/2)
			for j := range u16 {
				u16[j] = uint16(payload[j*2]) | uint16(payload[j*2+1])<<8
			}
			if s := cleanPPTText(decodeUTF16(u16)); s != "" {
				*out = append(*out, s)
			}
		}
	}
}

// decodeTextBytes 解码 TextBytesAtom 的负载。
//
// 按规范，该记录中每个字节对应一个 Latin-1 码点。但实际文件中偶有生成器
// 在此存放 UTF-8，若一律按 Latin-1 处理，中文会变成乱码。因此先检测负载
// 是否为合法的多字节 UTF-8：是则直接采用，否则按规范逐字节映射。
func decodeTextBytes(payload []byte) string {
	if isMultiByteUTF8(payload) {
		return strings.ReplaceAll(string(payload), "\x00", "")
	}

	var sb strings.Builder
	for _, b := range payload {
		if b != 0 {
			sb.WriteRune(rune(b))
		}
	}
	return sb.String()
}

// isMultiByteUTF8 判断字节序列是否为含多字节字符的合法 UTF-8。
//
// 纯 ASCII 不计入其中：它在两种解码方式下结果相同，无需特殊处理。
func isMultiByteUTF8(b []byte) bool {
	if !utf8.Valid(b) {
		return false
	}
	for _, c := range b {
		if c >= 0x80 {
			return true
		}
	}
	return false
}

// cleanPPTText 规范化提取到的文本。
//
// PowerPoint 使用 \r 表示段落分隔，用 \v 表示软换行，
// 并以 \x0b 等控制字符标记特殊内容，这些都需要转换或剔除。
func cleanPPTText(s string) string {
	s = strings.ReplaceAll(s, "\r", "\n")
	s = strings.ReplaceAll(s, "\v", "\n")

	var sb strings.Builder
	for _, r := range s {
		// 保留换行与制表符，剔除其余控制字符。
		if r == '\n' || r == '\t' || r >= 0x20 {
			sb.WriteRune(r)
		}
	}
	return strings.TrimSpace(sb.String())
}

// decodeUTF16 将 UTF-16 码元序列解码为字符串。
func decodeUTF16(u []uint16) string {
	runes := make([]rune, 0, len(u))
	for i := 0; i < len(u); i++ {
		c := u[i]
		// 处理代理对。
		if c >= 0xD800 && c <= 0xDBFF && i+1 < len(u) {
			lo := u[i+1]
			if lo >= 0xDC00 && lo <= 0xDFFF {
				runes = append(runes, ((rune(c)-0xD800)<<10|(rune(lo)-0xDC00))+0x10000)
				i++
				continue
			}
		}
		runes = append(runes, rune(c))
	}
	return string(runes)
}

// Texts 返回提取到的全部文本片段，按文档顺序排列。
func (d *LegacyDeck) Texts() []string {
	out := make([]string, len(d.texts))
	copy(out, d.texts)
	return out
}

// Text 返回演示文稿的全部文本，片段之间以换行分隔。
func (d *LegacyDeck) Text() (string, error) {
	return strings.Join(d.texts, "\n"), nil
}

// Save 始终返回 ErrReadOnlyFormat。
//
// 该方法的存在是为了让"尝试保存 .ppt"这一行为明确失败。
// 若需要修改，请读出文本后用 New 创建 .pptx。
func (d *LegacyDeck) Save(path string) error {
	return fmt.Errorf("pptgo: 无法写入 .ppt: %w", officecore.ErrReadOnlyFormat)
}

// byteReaderAt 让字节切片满足 io.ReaderAt 接口。
type byteReaderAt struct{ b []byte }

func newByteReaderAt(b []byte) *byteReaderAt { return &byteReaderAt{b} }

func (r *byteReaderAt) ReadAt(p []byte, off int64) (int, error) {
	if off < 0 || off >= int64(len(r.b)) {
		return 0, fmt.Errorf("offset %d out of range", off)
	}
	n := copy(p, r.b[off:])
	if n < len(p) {
		return n, fmt.Errorf("EOF")
	}
	return n, nil
}
