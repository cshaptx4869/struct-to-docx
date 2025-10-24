# struct-to-docx

基于 [docx](https://github.com/dolanmiu/docx) 库的结构描述生成 .docx 文件引擎。支持浏览器和 node.js 环境下使用。

```vue
<script setup lang="ts">
import { DocxBuilder } from "struct-to-docx"

function generateDocx() {
  const builder = new DocxBuilder()
  const doc = builder.addSection({
    children: [
      {
        type: "paragraph",
        options: {
          alignment: "center",
          children: [
            {
              type: "text",
              options: {
                text: " This is bold and red text.",
                bold: true,
                color: "FF0000",
                font: "Arial",
                size: DocxBuilder.textSize(12)
              }
            }
          ]
        }
      }
    ]
  }).render()
  builder.fileSave(doc, "example.docx")
}
</script>

<template>
  <div>
    <button @click="generateDocx">
      Generate DOCX
    </button>
  </div>
</template>

<style lang="scss" scoped></style>
```

> 详细示例参考 test 目录下的 [index.html](https://github.com/cshaptx4869/struct-to-docx/blob/main/test/index.html)
