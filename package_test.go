package pptgo

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"strings"
	"testing"
)

// readPackage 生成一个演示文稿并返回其全部部件，供结构断言使用。
func readPackage(t *testing.T, p *Presentation) map[string][]byte {
	t.Helper()

	data, err := p.Bytes()
	if err != nil {
		t.Fatalf("Bytes: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("生成的文件不是有效的 ZIP: %v", err)
	}

	parts := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("打开部件 %s: %v", f.Name, err)
		}
		b, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("读取部件 %s: %v", f.Name, err)
		}
		parts[f.Name] = b
	}
	return parts
}

// TestRequiredPartsPresent 验证规范强制要求的骨架部件均已写入。
// 缺少其中任何一个，PowerPoint 都会判定文件损坏。
func TestRequiredPartsPresent(t *testing.T) {
	p := New()
	p.AddSlide().SetTitleText("测试")

	parts := readPackage(t, p)

	required := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
		"ppt/slideMasters/slideMaster1.xml",
		"ppt/slideMasters/_rels/slideMaster1.xml.rels",
		"ppt/slideLayouts/slideLayout1.xml",
		"ppt/slideLayouts/_rels/slideLayout1.xml.rels",
		"ppt/theme/theme1.xml",
		"ppt/slides/slide1.xml",
		"ppt/slides/_rels/slide1.xml.rels",
	}

	for _, r := range required {
		if _, ok := parts[r]; !ok {
			t.Errorf("缺少必需部件: %s", r)
		}
	}
}

// TestEveryPartIsWellFormedXML 保证不会产出 XML 语法错误的部件。
func TestEveryPartIsWellFormedXML(t *testing.T) {
	p := New()
	p.AddSlide().SetTitleText("标题 & <特殊字符>").AddBullet("要点")

	for name, data := range readPackage(t, p) {
		if !strings.HasSuffix(name, ".xml") && !strings.HasSuffix(name, ".rels") {
			continue
		}
		dec := xml.NewDecoder(bytes.NewReader(data))
		for {
			_, err := dec.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				t.Errorf("部件 %s XML 格式错误: %v", name, err)
				break
			}
		}
	}
}

// TestEverySlideDeclaredInContentTypes 覆盖此前的缺陷：
// 幻灯片部件未在内容类型清单中声明，PowerPoint 会因此拒绝打开文件。
func TestEverySlideDeclaredInContentTypes(t *testing.T) {
	p := New()
	for i := 0; i < 3; i++ {
		p.AddSlide().SetTitleText(fmt.Sprintf("第 %d 页", i+1))
	}

	parts := readPackage(t, p)
	ct := string(parts["[Content_Types].xml"])

	for i := 1; i <= 3; i++ {
		want := fmt.Sprintf("/ppt/slides/slide%d.xml", i)
		if !strings.Contains(ct, want) {
			t.Errorf("[Content_Types].xml 未声明 %s", want)
		}
	}
}

// TestNoDanglingRelationships 是最关键的一项检查：
// 关系若指向不存在的部件，即为悬空引用，会直接导致文件损坏。
func TestNoDanglingRelationships(t *testing.T) {
	p := New()
	s := p.AddSlide().SetTitleText("含图片")
	s.AddImageBytes([]byte{0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A}, 100, 100, 200, 200)

	p.AddSlide().SetTitleText("无图片")

	parts := readPackage(t, p)

	type rels struct {
		Rel []struct {
			ID         string `xml:"Id,attr"`
			Target     string `xml:"Target,attr"`
			TargetMode string `xml:"TargetMode,attr"`
		} `xml:"Relationship"`
	}

	for name, data := range parts {
		if !strings.HasSuffix(name, ".rels") {
			continue
		}

		var r rels
		if err := xml.Unmarshal(data, &r); err != nil {
			t.Errorf("解析 %s: %v", name, err)
			continue
		}

		// 关系中的 Target 相对于所属部件的所在目录。
		baseDir := path.Dir(path.Dir(name))

		for _, rel := range r.Rel {
			if rel.TargetMode == "External" || rel.Target == "" {
				continue
			}
			resolved := rel.Target
			if !strings.HasPrefix(resolved, "/") {
				resolved = path.Join(baseDir, rel.Target)
			} else {
				resolved = strings.TrimPrefix(resolved, "/")
			}

			if _, ok := parts[resolved]; !ok {
				t.Errorf("%s: 关系 %s 指向不存在的部件 %s（原始 Target=%s）",
					name, rel.ID, resolved, rel.Target)
			}
		}
	}
}

// TestSlideRelsAlwaysReferenceLayout 验证即使幻灯片没有图片，
// 也必须建立与版式的关联。
func TestSlideRelsAlwaysReferenceLayout(t *testing.T) {
	p := New()
	p.AddSlide().SetTitleText("无图片")

	parts := readPackage(t, p)
	relsData, ok := parts["ppt/slides/_rels/slide1.xml.rels"]
	if !ok {
		t.Fatal("无图片的幻灯片缺少关系部件")
	}
	if !strings.Contains(string(relsData), "slideLayout1.xml") {
		t.Error("幻灯片关系中缺少版式引用")
	}
}

// TestPresentationHasSlideSizeAndMaster 验证 presentation.xml 中
// 两个规范必填项均已写入。
func TestPresentationHasSlideSizeAndMaster(t *testing.T) {
	p := New()
	p.AddSlide().SetTitleText("测试")

	parts := readPackage(t, p)
	pres := string(parts["ppt/presentation.xml"])

	if !strings.Contains(pres, "<p:sldSz") {
		t.Error("presentation.xml 缺少 sldSz（幻灯片尺寸为必填项）")
	}
	if !strings.Contains(pres, "sldMasterIdLst") {
		t.Error("presentation.xml 缺少 sldMasterIdLst")
	}
}

// TestImageNumberingIsConsistent 覆盖此前的编号缺陷：
// 图片部件的写入编号与关系中的引用编号必须一致。
func TestImageNumberingIsConsistent(t *testing.T) {
	p := New()
	s1 := p.AddSlide().SetTitleText("第一页")
	s1.AddImageBytes([]byte("fake-png-1"), 0, 0, 100, 100)
	s1.AddImageBytes([]byte("fake-png-2"), 0, 0, 100, 100)

	s2 := p.AddSlide().SetTitleText("第二页")
	s2.AddImageBytes([]byte("fake-png-3"), 0, 0, 100, 100)

	parts := readPackage(t, p)

	var media []string
	for name := range parts {
		if strings.HasPrefix(name, "ppt/media/") {
			media = append(media, name)
		}
	}
	if len(media) != 3 {
		t.Errorf("期望写入 3 个图片部件，实际 %d 个: %v", len(media), media)
	}

	// 每一条图片关系都应能对应到真实存在的部件，
	// 该断言由 TestNoDanglingRelationships 覆盖；此处额外确认编号规则本身。
	for slideNum, count := range map[int]int{1: 2, 2: 1} {
		for i := 0; i < count; i++ {
			want := fmt.Sprintf("ppt/media/image%d.png", imageNumber(slideNum, i))
			if _, ok := parts[want]; !ok {
				t.Errorf("缺少图片部件 %s", want)
			}
		}
	}
}
