import fs from "fs";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

export default async function handler(req, res) {
    if (req.method !== "POST") return res.status(405).end();

    const data = req.body;
    const content = fs.readFileSync("template.docx", "binary");

    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip);

    doc.render({
        NGAY: data.ngay,
        BEN_A: data.benA,
        BEN_B: data.benB,
        BEN_C: data.benC,
        TEN_HANG: data.hang,
        SO_LUONG: data.soluong,
        GIA: data.gia
    });

    const buf = doc.getZip().generate({ type: "nodebuffer" });

    res.setHeader("Content-Disposition", "attachment; filename=hopdong.docx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buf);
}
