import fs from "node:fs"
import path from "node:path"
import url from "node:url"
// eslint-disable-next-line antfu/no-import-dist
import { DocxBuilder, Packer } from "../dist/index.esm.js"

const __dirname = path.dirname(url.fileURLToPath(import.meta.url))
const output = path.join(__dirname, "output", `${Date.now()}.docx`)
const template = JSON.parse(
  fs.readFileSync(path.join(__dirname, "./assets/template.json"), "utf-8"),
)
const data = {
  "inspect.0.level": "S-1",
  "inspect.0.sampleQuantity": "10000",
  "inspect.0.aql": "2.5",
  "inspect.0.acRe": "0/1",
  "inspect.0.defectiveQty": "0",
  "inspect.0.result": "合格\nPASS",
  "inspect.1.level": "S-1",
  "inspect.1.sampleQuantity": "3",
  "inspect.1.aql": "2.5",
  "inspect.1.acRe": "0/1",
  "inspect.1.defectiveQty": "0",
  "inspect.1.result": "合格\nPASS",
  "inspect.2.level": "S-1",
  "inspect.2.sampleQuantity": "10000",
  "inspect.2.aql": "0.10",
  "inspect.2.acRe": "0/1",
  "inspect.2.defectiveQty": "0",
  "inspect.2.result": "合格\nPASS",
  "inspect.3.reportNameAndNo": "填写报告编号和名称\nReport No. & Name",
  "inspect.3.defectiveQty": "0",
  "inspect.3.result": "合格\nPASS",
  "inspect.4.level": "S-1",
  "inspect.4.sampleQuantity": "10000",
  "inspect.4.aql": "2.5",
  "inspect.4.acRe": "0/1",
  "inspect.4.defectiveQty": "0",
  "inspect.4.point": "100",
  "inspect.4.result": "合格\nPASS",
  "recordNumber": "recordNumber",
  "versionNumber": "versionNumber",
  "formatApprover": "formatApprover",
  "approvalDate": "approvalDate",
  "title": "title",
  "subTitle": "subTitle",
  "reportNo": "reportNo",
  "inspectionDate": "inspectionDate",
  "name": "name",
  "itemNo": "itemNo",
  "lotNo": "lotNo",
  "supplier": "supplier",
  "purchasingNo": "purchasingNo",
  "quantity": "quantity",
  "inspectionStandards": "inspectionStandards",
  "inspector": "inspector",
  "approver": "approver",
  "testConclusion": "testConclusion",
  "remark": "●remark",
}

// export docx
const docxBuilder = new DocxBuilder()
const doc = docxBuilder
  .setProperties({
    title: "QC Report",
    creator: "struct-to-docx",
  })
  .setDefaultFont("Arial")
  .setDefaultFontSize(8)
  .addSection(template)
  .render(data)

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(output, buffer)
  console.log("🚀 ~ output:", output)
})
