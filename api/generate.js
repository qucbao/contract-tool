import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

export default async function handler(req, res) {
    try {
        if (req.method !== "POST") {
            return res.status(405).json({ error: "Method not allowed" });
        }

        const templatePath = path.join(process.cwd(), "template.docx");

        if (!fs.existsSync(templatePath)) {
            return res.status(500).json({
                error: "Template not found",
                path: templatePath
            });
        }

        const content = fs.readFileSync(templatePath, "binary");

        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true
        });

        doc.render(req.body);

        const buf = doc.getZip().generate({ type: "nodebuffer" });

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        );
        res.setHeader(
            "Content-Disposition",
            "attachment; filename=hopdong.docx"
        );

        return res.status(200).send(buf);
    } catch (err) {
        console.error("DOCX ERROR:", err);
        return res.status(500).json({
            error: "DOCX_GENERATE_FAILED",
            message: err.message
        });
    }
}
