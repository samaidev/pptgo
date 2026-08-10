# pptgo

简洁的 Go 库，用于读写 .pptx 演示文稿。设计参考 python-pptx，自研 OOXML 渲染。

## 安装

```bash
go get github.com/samaidev/pptgo
```

## 快速开始

```go
package main

import "github.com/samaidev/pptgo"

func main() {
    prs := pptgo.New()
    prs.SetTitle("产品介绍").SetAuthor("SamAI")

    s1 := prs.AddSlide()
    s1.SetLayout(pptgo.LayoutTitle)
    s1.AddTitle("GoOffice")
    s1.AddSubtitle("跨平台办公套件")

    s2 := prs.AddSlide()
    s2.AddTitle("核心功能")
    s2.AddBullets([]string{"文档", "表格", "演示"})

    prs.Save("presentation.pptx")
}
```

## 高级功能

- 3 种布局：LayoutTitle / LayoutContent / LayoutBlank
- 过渡效果：`slide.SetTransition("fade", 500)`
- 元素动画：`slide.AddAnimation("title", "fly", 0, 500)`
- 演讲者备注：`slide.SetNotes("备注内容")`
- 背景色：`slide.SetBgColor("#1e293b")`

## 保真编辑（重要）

打开已存在的 `.pptx` 时，`pptgo` 同样采用**保真编辑**策略：未修改的幻灯片母版、版式、主题与图片原样写回，仅替换目标文本节点。

```go
deck, _ := pptgo.OpenPPTX("deck.pptx")
n := deck.ReplaceText("旧标题", "新标题")
deck.SaveAs("deck-new.pptx")
```

## 旧版 .ppt 支持（只读）

旧版二进制 `.ppt` 为 OLE2/CFB 复合文档，支持**只读**读取：

```go
d, _ := pptgo.OpenLegacy("old.ppt")
for i := 0; i < d.SlideCount(); i++ {
    text, _ := d.SlideText(i)
    fmt.Println(text)
}
```

> 设计约束：旧版 Office 文件仅支持读取，编辑请使用现代 `.pptx` 格式。

## 许可证

MIT
