import {
	Document,
	Paragraph,
	Packer,
	TextRun,
	FrameAnchorType,
	HorizontalPositionAlign,
	VerticalPositionAlign,
	PageOrientation,
	AlignmentType,
    Table,
    TableCell, TableRow, WidthType,
} from "docx";
import { saveAs } from "file-saver";


const table = new Table({
    columnWidths: [12500],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 12500,
                        type: WidthType.DXA,
                    },
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "1. FROM (IncludeZipCode)",
                                    size: 6,
                                }),
                            ]
                        }),
                    ],
                }),
            ]
        })
    ]
})



const generateDoc = () => {
	const doc = new Document({
		sections: [
			{
				properties: {
					page: {
						size: {
							orientation: PageOrientation.LANDSCAPE,
						},
					},
				},
				children: [
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 0,
								y: -900,
							},
							width: 10000,
							height: 300,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.TOP,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							bottom: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							left: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							right: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
						},
						children: [
							new TextRun({
								text: "REQUISITION AND INVOICE / SHIPPING DOCUMENT",
								bold: true,
							}),
						],
					}),
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 10115,
								y: -900,
							},
							width: 3700,
							height: 300,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.TOP,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							bottom: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							left: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							right: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
						},
						children: [
							new TextRun({
								text: "FormApprovedOMBNo. 0704-0246",
								italics: true,
                                size: 16,
							}),
						],
					}),
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: -610,
							},
							width: 13810,
							height: 800,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.TOP,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							bottom: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							left: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							right: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
						},
						children: [
							new TextRun({
								text: "Public reporting burden for this collection of information is estimated to average 1 hour per response, including the time for reviewing instructions, searching existing data sources, gathering and maintaining the data needed, and completing and reviewing the collection of information. Send comments regarding this burden estimate or any other aspect of this collection of information, including suggestions for reducing this burden, to Washington Headquarters Services, Directorate for Information Operations and Reports, 1215 Jefferson Davis Highway, Suite 1204, Arlington, VA 22202-4302, and to the Office of Management and Budget, Paperwork Reduction Project (0704-0246), Washington, DC20503.",
                                size: 12,
							}),
						],
					}),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        frame: {
							position: {
								x: 0,
								y: -100,
							},
							width: 13810,
							height: 300,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.TOP,
							},
						}, 
                        children: [
                            new TextRun({
                                text: "Please do not return your completed form to either of these addresses. Return completed form to the address in item 2.",
                                size: 14,
                                allCaps: true,
                            })
                        ]
                    }),
                    new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 170,
							},
							width: 8000,
							height: 1000,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.TOP,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							bottom: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							left: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
							right: {
								color: "auto",
								space: 1,
								value: "single",
								size: 6,
							},
						},
						children: [
							new TextRun({
								text: "1. FROM (IncludeZIPCode)",
                                size: 10,
							}),
						],
					}),
                    new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 350,
							},
							width: 5500,
							height: 1000,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.TOP,
							},
						},
                        children: [
							new TextRun({
								text: "TRIDENT REFIT FACILITY, CODE 432",
                                size: 18,
							}),
                            new TextRun({
                                text: "POC: LEE SAVELSON, PHONE (912) 573-3881/ FAX (912) 573-3709",
                                size: 18,
                                break: 1,
                            }),
                            new TextRun({
                                text: "KINGS BAY GA 31547-6000",
                                size: 18,
                                break: 1,
                            }),
						],
					}),
				],
			},
		],
        children: [
            table,
        ]
	});

	Packer.toBlob(doc).then((blob) => {
		saveAs(blob, "Test Document");
	});
};

export default generateDoc;
