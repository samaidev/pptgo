package pptgo

import (
	"strings"
	"testing"
)

// TestReadRealPptx 回归：对真实 PowerPoint 2007+ 文档（来自 LibreOffice 测试数据）
// 应成功解析幻灯片数量并提取文本与备注。
func TestReadRealPptx(t *testing.T) {
	const path = "testdata/sample.pptx"
	d, err := OpenPPTX(path)
	if err != nil {
		t.Skipf("跳过：未找到真实文件 %s（%v）", path, err)
	}
	if d.SlideCount() == 0 {
		t.Fatalf("真实 pptx 未解析出任何幻灯片")
	}
	txt, err := d.Text()
	if err != nil {
		t.Fatalf("Text() 错误: %v", err)
	}
	if strings.TrimSpace(txt) == "" {
		t.Fatalf("真实 pptx 提取文本为空")
	}
}

// TestReadRealLegacyPpt 回归：对真实 OLE2(.ppt) 文档应提取出大纲文本。
func TestReadRealLegacyPpt(t *testing.T) {
	const path = "testdata/sample.ppt"
	d, err := OpenLegacy(path)
	if err != nil {
		t.Skipf("跳过：未找到真实文件 %s（%v）", path, err)
	}
	txt, _ := d.Text()
	if !strings.Contains(txt, "Click to edit") {
		t.Fatalf("真实 ppt 未提取出预期大纲文本，preview=%q", txt[:min(60, len(txt))])
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
