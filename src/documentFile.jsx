import { Document, Paragraph, Packer, TextRun } from 'docx';
import { saveAs } from "file-saver";

const generateDoc= () => {
    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Testing First Paragraph"),
                        new TextRun("Testing Second Line"),
                        new TextRun({
                            text: "Spicy Text",
                            bold: true,
                        }),
                        new TextRun({
                            text: "Testing Text block",
                            bold: true,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Testing First Paragraph"),
                        new TextRun("Testing Second Line"),
                        new TextRun({
                            text: "Spicy Text",
                            bold: true,
                        }),
                        new TextRun({
                            text: "Testing Text block",
                            bold: true,
                        }),
                    ],
                }),
            ],
        }],
    });

    Packer.toBlob(doc).then((blob) => {
        saveAs(blob, "Test Document");
    });
};

// Packer.toBuffer(doc).then((buffer) => {
//     fs.writeFileSync("Test Doc.docx", buffer);
// });

export default generateDoc;