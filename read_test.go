package pptgo

import (
	"bytes"
	"encoding/binary"
	"errors"
	"strings"
	"testing"

	"github.com/samaidev/officecore"
)

// TestOpenPPTXRoundTrip 验证本库生成的文件能被自身读回，
// 这是读写两条路径一致性的基本检验。
func TestOpenPPTXRoundTrip(t *testing.T) {
	p := New()
	p.AddSlide().SetTitleText("第一页标题").AddBullet("要点一").AddBullet("要点二")
	p.AddSlide().SetTitleText("第二页标题")

	data, err := p.Bytes()
	if err != nil {
		t.Fatalf("生成 pptx: %v", err)
	}

	deck, err := OpenPPTXBytes(data)
	if err != nil {
		t.Fatalf("OpenPPTXBytes: %v", err)
	}

	if deck.SlideCount() != 2 {
		t.Fatalf("期望 2 张幻灯片，实际 %d 张", deck.SlideCount())
	}

	first, err := deck.SlideText(0)
	if err != nil {
		t.Fatalf("SlideText: %v", err)
	}
	for _, want := range []string{"第一页标题", "要点一", "要点二"} {
		if !strings.Contains(first, want) {
			t.Errorf("第一页缺少文本 %q，实际内容:\n%s", want, first)
		}
	}
}

// TestSlideNotesExtracted 验证演讲者备注（notesSlide）能被读取，
// 符合 LibreOffice 对 notesSlide 的处理范式。
func TestSlideNotesExtracted(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.SetTitleText("标题页面")
	s.SetNotes("这是演讲者备注，不应出现在幻灯片正文里。")

	data, err := p.Bytes()
	if err != nil {
		t.Fatalf("生成 pptx: %v", err)
	}
	deck, err := OpenPPTXBytes(data)
	if err != nil {
		t.Fatalf("OpenPPTXBytes: %v", err)
	}
	if deck.SlideCount() != 1 {
		t.Fatalf("期望 1 张幻灯片，实际 %d", deck.SlideCount())
	}
	body, err := deck.SlideText(0)
	if err != nil {
		t.Fatalf("SlideText: %v", err)
	}
	if strings.Contains(body, "演讲者备注") {
		t.Errorf("备注文本不应出现在正文里:\n%s", body)
	}
	notes, err := deck.SlideNotes(0)
	if err != nil {
		t.Fatalf("SlideNotes: %v", err)
	}
	if !strings.Contains(notes, "演讲者备注") {
		t.Errorf("SlideNotes 未提取到备注，实际: %q", notes)
	}
}

// TestSlideOrderIsNumeric 覆盖一个容易出错的细节：
// 按字典序排序会把 slide10 排在 slide2 前面。
func TestSlideOrderIsNumeric(t *testing.T) {
	p := New()
	for i := 1; i <= 12; i++ {
		p.AddSlide().SetTitleText(strings.Repeat("x", 1) + string(rune('0'+i%10)))
	}

	data, err := p.Bytes()
	if err != nil {
		t.Fatalf("生成 pptx: %v", err)
	}
	deck, err := OpenPPTXBytes(data)
	if err != nil {
		t.Fatalf("OpenPPTXBytes: %v", err)
	}

	if deck.SlideCount() != 12 {
		t.Fatalf("期望 12 张幻灯片，实际 %d 张", deck.SlideCount())
	}

	// 部件顺序应为 slide1..slide12，而非 slide1, slide10, slide11...
	for i, part := range deck.slideParts {
		if got := slideNumOf(part); got != i+1 {
			t.Errorf("第 %d 个部件为 %s（编号 %d），顺序有误", i, part, got)
		}
	}
}

// TestPPTXFidelity 验证修改文本后，其余部件保持不变。
func TestPPTXFidelity(t *testing.T) {
	p := New()
	s := p.AddSlide().SetTitleText("原始标题")
	s.AddImageBytes([]byte{0x89, 'P', 'N', 'G', 0xDE, 0xAD}, 0, 0, 100, 100)

	data, err := p.Bytes()
	if err != nil {
		t.Fatalf("生成 pptx: %v", err)
	}

	before := readPackage2(t, data)

	deck, err := OpenPPTXBytes(data)
	if err != nil {
		t.Fatalf("OpenPPTXBytes: %v", err)
	}
	if n := deck.ReplaceText("原始标题", "新标题"); n != 1 {
		t.Fatalf("期望替换 1 处，实际 %d 处", n)
	}

	saved, err := deck.SaveBytes()
	if err != nil {
		t.Fatalf("SaveBytes: %v", err)
	}
	after := readPackage2(t, saved)

	// 母版、版式、主题、图片都不应发生变化。
	for _, part := range []string{
		"ppt/slideMasters/slideMaster1.xml",
		"ppt/slideLayouts/slideLayout1.xml",
		"ppt/theme/theme1.xml",
		"ppt/media/image101.png",
	} {
		if !bytes.Equal(before[part], after[part]) {
			t.Errorf("部件 %s 在编辑文本后发生了变化", part)
		}
	}

	// 编辑本身必须生效。
	reopened, err := OpenPPTXBytes(saved)
	if err != nil {
		t.Fatalf("重新打开: %v", err)
	}
	text, _ := reopened.SlideText(0)
	if !strings.Contains(text, "新标题") {
		t.Errorf("替换未生效: %q", text)
	}
}

// readPackage2 读取 pptx 字节流中的全部部件。
func readPackage2(t *testing.T, data []byte) map[string][]byte {
	t.Helper()

	deck, err := OpenPPTXBytes(data)
	if err != nil {
		t.Fatalf("读取包: %v", err)
	}

	out := make(map[string][]byte)
	for _, name := range deck.pkg.Names() {
		out[name] = deck.pkg.Part(name)
	}
	return out
}

func TestSlideTextRejectsOutOfRange(t *testing.T) {
	p := New()
	p.AddSlide().SetTitleText("唯一一页")

	data, _ := p.Bytes()
	deck, err := OpenPPTXBytes(data)
	if err != nil {
		t.Fatalf("OpenPPTXBytes: %v", err)
	}

	if _, err := deck.SlideText(5); err == nil {
		t.Error("越界下标应返回错误")
	}
	if _, err := deck.SlideText(-1); err == nil {
		t.Error("负数下标应返回错误")
	}
}

func TestOpenPPTXRejectsNonPresentation(t *testing.T) {
	if _, err := OpenPPTXBytes([]byte("这不是一个 pptx")); err == nil {
		t.Error("非 ZIP 数据应返回错误")
	}
}

// === .ppt 旧版格式测试 ===

// pptRecord 构造一条 PowerPoint 记录：
// 2 字节版本与实例号、2 字节类型、4 字节长度，随后是负载。
func pptRecord(verInst, typ uint16, payload []byte) []byte {
	hdr := make([]byte, 8)
	binary.LittleEndian.PutUint16(hdr[0:], verInst)
	binary.LittleEndian.PutUint16(hdr[2:], typ)
	binary.LittleEndian.PutUint32(hdr[4:], uint32(len(payload)))
	return append(hdr, payload...)
}

// utf16LE 将字符串编码为 UTF-16LE 字节序列。
func utf16LE(s string) []byte {
	var out []byte
	for _, r := range s {
		out = append(out, byte(r), byte(r>>8))
	}
	return out
}

func TestExtractPPTTextBytesAtom(t *testing.T) {
	stream := pptRecord(0x00, recTextBytesAtom, []byte("Hello World"))

	got := extractPPTText(stream)
	if len(got) != 1 || got[0] != "Hello World" {
		t.Errorf("期望 [Hello World]，实际 %q", got)
	}
}

// TestExtractPPTTextCharsAtom 验证 UTF-16 文本，中文必须正确解码。
func TestExtractPPTTextCharsAtom(t *testing.T) {
	stream := pptRecord(0x00, recTextCharsAtom, utf16LE("季度总结"))

	got := extractPPTText(stream)
	if len(got) != 1 || got[0] != "季度总结" {
		t.Errorf("期望 [季度总结]，实际 %q", got)
	}
}

// TestExtractPPTTextRecursesIntoContainers 验证容器记录被递归展开。
// 真实的 .ppt 文件中，文本原子总是嵌套在多层容器之内。
//
// 此处使用 ASCII 文本，因为 TextBytesAtom 按规范只能存放单字节字符；
// 非 ASCII 内容会以 TextCharsAtom 的 UTF-16 形式存储。
func TestExtractPPTTextRecursesIntoContainers(t *testing.T) {
	inner := pptRecord(0x00, recTextBytesAtom, []byte("Nested text"))
	// verInst 低 4 位为 0xF 表示容器。
	container := pptRecord(0x0F, 0x03E8, inner)
	outer := pptRecord(0x0F, 0x03E8, container)

	got := extractPPTText(outer)
	if len(got) != 1 || got[0] != "Nested text" {
		t.Errorf("容器中的文本未被提取，实际 %q", got)
	}
}

// TestExtractPPTTextRecursesWithUnicode 用 UTF-16 原子验证同样的递归路径，
// 这对应真实文件中存放中文的方式。
func TestExtractPPTTextRecursesWithUnicode(t *testing.T) {
	inner := pptRecord(0x00, recTextCharsAtom, utf16LE("嵌套文本"))
	container := pptRecord(0x0F, 0x03E8, inner)

	got := extractPPTText(container)
	if len(got) != 1 || got[0] != "嵌套文本" {
		t.Errorf("容器中的 Unicode 文本未被提取，实际 %q", got)
	}
}

// TestExtractPPTTextHandlesTruncation 确保异常文件不会导致越界。
func TestExtractPPTTextHandlesTruncation(t *testing.T) {
	stream := pptRecord(0x00, recTextBytesAtom, []byte("正常"))
	// 追加一个声称长度极大、实际数据不足的记录头。
	bad := make([]byte, 8)
	binary.LittleEndian.PutUint16(bad[2:], recTextBytesAtom)
	binary.LittleEndian.PutUint32(bad[4:], 0xFFFF)
	stream = append(stream, bad...)

	got := extractPPTText(stream) // 不应 panic
	if len(got) != 1 {
		t.Errorf("截断记录应被忽略，实际提取出 %q", got)
	}
}

// TestExtractPPTTextGuardsRecursionDepth 防止构造异常的文件耗尽栈空间。
func TestExtractPPTTextGuardsRecursionDepth(t *testing.T) {
	// 构造深度嵌套的容器。
	data := pptRecord(0x00, recTextBytesAtom, []byte("底层"))
	for i := 0; i < 50; i++ {
		data = pptRecord(0x0F, 0x03E8, data)
	}

	// 超过深度限制后停止即可，重点是不能崩溃。
	_ = extractPPTText(data)
}

func TestCleanPPTText(t *testing.T) {
	cases := map[string]string{
		"正常文本":         "正常文本",
		"带\r回车":        "带\n回车",
		"带\v软换行":       "带\n软换行",
		"带\x01控制字符":    "带控制字符",
		"  首尾空白  ":     "首尾空白",
	}
	for in, want := range cases {
		if got := cleanPPTText(in); got != want {
			t.Errorf("cleanPPTText(%q) = %q，期望 %q", in, got, want)
		}
	}
}

func TestDecodeUTF16SurrogatePair(t *testing.T) {
	// U+1F600 的代理对表示。
	got := decodeUTF16([]uint16{0xD83D, 0xDE00})
	if got != "\U0001F600" {
		t.Errorf("代理对解码错误: %q", got)
	}
}

// TestLegacySaveReturnsReadOnlyError 验证尝试写入 .ppt 会明确失败。
func TestLegacySaveReturnsReadOnlyError(t *testing.T) {
	d := &LegacyDeck{}

	err := d.Save("out.ppt")
	if err == nil {
		t.Fatal("写入 .ppt 应当返回错误")
	}
	if !errors.Is(err, officecore.ErrReadOnlyFormat) {
		t.Errorf("错误应可通过 errors.Is 匹配 ErrReadOnlyFormat，实际为: %v", err)
	}
}

func TestLegacyDeckText(t *testing.T) {
	d := &LegacyDeck{texts: []string{"第一段", "第二段"}}

	text, err := d.Text()
	if err != nil {
		t.Fatalf("Text: %v", err)
	}
	if text != "第一段\n第二段" {
		t.Errorf("Text = %q", text)
	}

	if len(d.Texts()) != 2 {
		t.Errorf("Texts 应返回 2 个片段")
	}
}

func TestOpenLegacyRejectsNonCFB(t *testing.T) {
	if _, err := OpenLegacyBytes([]byte("PK\x03\x04 这是一个 zip")); err == nil {
		t.Error("将 pptx 作为 ppt 打开应返回错误")
	}
}
