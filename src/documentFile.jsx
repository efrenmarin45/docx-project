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
	TableCell,
	TableRow,
	WidthType,
	TextDirection,
	TDirection,
	ShadingType,
} from "docx";
import { saveAs } from "file-saver";

const table = new Table({
	columnWidths: [250],
	rows: [
		new TableRow({
			children: [
				new TableCell({
					width: {
						size: 250,
						type: WidthType.DXA,
					},
					children: [new Paragraph({})],
				}),
				new TableCell({
					width: {
						size: 250,
						type: WidthType.DXA,
					},
					children: [new Paragraph({})],
				}),
			],
		}),
	],
});

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
					//Header and Title
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
					//OMB Number
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 10115,
								y: -900,
							},
							width: 3695,
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
								text: "TEST TEXT",
								italics: true,
								size: 16,
							}),
						],
					}),
					//Disclosure Main Text
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: -610,
							},
							width: 13805,
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
								text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Dictumst vestibulum rhoncus est pellentesque elit ullamcorper dignissim cras. Ipsum dolor sit amet consectetur adipiscing elit pellentesque habitant. Nunc sed augue lacus viverra vitae congue eu consequat ac. Eget aliquet nibh praesent tristique magna sit amet purus. Neque laoreet suspendisse interdum consectetur libero id faucibus. In pellentesque massa placerat duis ultricies. Dui vivamus arcu felis bibendum ut tristique. Molestie ac feugiat sed lectus. Urna cursus eget nunc scelerisque viverra mauris in. Nisl tincidunt eget nullam non. At risus viverra adipiscing.",
								size: 12,
							}),
						],
					}),
					//Disclosure Sub Text
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
							}),
						],
					}),
					//FROM Row #1 Address Header
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 165,
							},
							width: 8000,
							height: 1015,
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
					//FROM Row Address
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 450,
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
								text: "EXAMPLE ADDRESS",
								size: 18,
							}),
							new TextRun({
								text: "POC: BOB BUILDER, PHONE (222) 222-2222/ FAX (111) 111-1111",
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
					//FROM Row Cell - Sheet No.
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 8110,
								y: 165,
							},
							width: 1010,
							height: 400,
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
								text: "Sheet No.",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "1",
								break: 1,
								size: 16,
							}),
						],
					}),
					//FROM Row Cell - No. of Sheets
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9230,
								y: 165,
							},
							width: 1010,
							height: 400,
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
								text: "No. of Sheets",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "1",
								break: 1,
								size: 16,
							}),
						],
					}),
					//FROM Row Cell - #5 Requisition Date
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 10350,
								y: 165,
							},
							width: 1180,
							height: 400,
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
								text: "5. Requisition Date",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "01/01/2021",
								break: 1,
								size: 16,
							}),
						],
					}),
					//FROM Row Cell - #6 Requisition Number
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 11640,
								y: 165,
							},
							width: 2170,
							height: 400,
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
								text: "6. Requisition No",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "N44466-0000",
								break: 1,
								size: 16,
							}),
						],
					}),
					//FROM Row Cell - #7 Date/Material Required
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8110,
								y: 540,
							},
							width: 3420,
							height: 635,
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
								text: "7. Date/Material Required (YYMMDD)",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "\t210101",
								break: 2,
								size: 18,
							}),
						],
					}),
					//FROM Row Cell - #8 Priority
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 11640,
								y: 540,
							},
							width: 2170,
							height: 640,
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
								text: "8. Priority",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "Yes",
								break: 2,
								size: 18,
							}),
						],
					}),
					//TO Row #2 Address Header
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 1160,
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
								text: "2. TO: (IncludeZIPCode)",
								size: 10,
							}),
						],
					}),
					//TO Row Address
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 1400,
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
								text: "TEST ADDRESS",
								size: 18,
							}),
							new TextRun({
								text: "BLDG 0000",
								size: 18,
								break: 1,
							}),
							new TextRun({
								text: "KINGS BAY, GA 31547-6000",
								size: 18,
								break: 1,
							}),
						],
					}),
					//TO Row Cell - #9 Authority or Purpose
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8110,
								y: 1165,
							},
							width: 5700,
							height: 510,
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
								text: "9. Authority or Purpose",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "TOOLS and Others",
								allCaps: true,
								break: 1,
								size: 18,
							}),
						],
					}),
					//TO Row Cell - #10 Signature
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8110,
								y: 1660,
							},
							width: 2780,
							height: 500,
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
								text: "10. Signature",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "Bob Builder",
								allCaps: true,
								break: 1,
								size: 18,
							}),
						],
					}),
					//TO Row Cell - 11a. Voucher Number and Date (YYMMDD)
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 11000,
								y: 1660,
							},
							width: 2810,
							height: 500,
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
								text: "11a. Voucher Number & Date (YYMMDD)",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "D123456 210101",
								allCaps: true,
								break: 1,
								size: 18,
							}),
						],
					}),
					//SHIP TO Row #3 Address Header
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 2135,
							},
							width: 8000,
							height: 1350,
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
								text: "3. SHIP TO - MARK FOR",
								size: 10,
							}),
						],
					}),
					//SHIP TO Row Address
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 2350,
							},
							width: 5500,
							height: 1800,
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
								text: "TEST ADDRESS",
								size: 18,
							}),
							new TextRun({
								text: "0000 MAIN ROAD",
								size: 18,
								break: 1,
							}),
							new TextRun({
								text: "ATLANTIS, NY 23502",
								size: 18,
								break: 1,
							}),
							new TextRun({
								text: "POC: BOB BUILDER (555) 555-5555 x000",
								size: 18,
								break: 1,
							}),
							new TextRun({
								text: "** E-MAIL TRACKING INFO TO BOB.BUILDER@WORLD.ALL",
								size: 18,
								break: 1,
							}),
						],
					}),
					//SHIP TO Row Cell - #12 Date Shipped (YYMMDD)
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8110,
								y: 2135,
							},
							width: 2780,
							height: 500,
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
								text: "12. Date Shipped",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "210101",
								allCaps: true,
								break: 1,
								size: 18,
							}),
						],
					}),
					//SHIP TO Row Cell - 11b EMPTY
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 11000,
								y: 2135,
							},
							width: 2810,
							height: 500,
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
								text: "11b.",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//SHIP TO Row Cell - #13 Mode of Shipment
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8110,
								y: 2615,
							},
							width: 2780,
							height: 500,
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
								text: "13. Mode of Shipment",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//SHIP TO Row Cell - #14 Bill of Landing Number
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 11000,
								y: 2615,
							},
							width: 2810,
							height: 500,
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
								text: "14. Bill of Lading Number",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//SHIP TO Row Cell - #15 Air Movement Designator or Port Reference Number
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8110,
								y: 3095,
							},
							width: 5700,
							height: 390,
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
								text: "15. Air Movement Designator or Port Reference No.",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Row 4 Cell - #4 Appropriations Symbol and Subhead
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 3470,
							},
							width: 3500,
							height: 650,
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
								text: "4. Appropriations Symbol and Subhead",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Row 4 Cell - Object Class
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 3600,
								y: 3470,
							},
							width: 550,
							height: 650,
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
								text: "Object Class",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Row 4 Cell - Bureau Control Co.
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 4255,
								y: 3470,
							},
							width: 980,
							height: 650,
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
								text: "Bureau Control No.",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Row 4 Cell - Subal-Lot
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 5335,
								y: 3470,
							},
							width: 550,
							height: 650,
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
								text: "Subal-Lot",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Row 4 Cell - Authorization ACCTG Activity
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 5990,
								y: 3470,
							},
							width: 2010,
							height: 650,
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
								text: "Authorization ACCTG Activity",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Row 4 Cell - Trans. Type
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 8100,
								y: 3470,
							},
							width: 1010,
							height: 650,
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
								text: "Trans Type",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Row 4 Cell - Property ACCTG Activity
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9215,
								y: 3470,
							},
							width: 1505,
							height: 650,
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
								text: "Property ACCTG Activity",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Row 4 Cell - Country
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 10825,
								y: 3470,
							},
							width: 570,
							height: 650,
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
								text: "Country",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Row 4 Cell - Cost Code
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 11500,
								y: 3470,
							},
							width: 1510,
							height: 650,
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
								text: "Cost Code",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "NLAF",
								allCaps: true,
								size: 18,
								break: 2,
							}),
						],
					}),
					//Row 4 Cell - Amount
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 13125,
								y: 3470,
							},
							width: 685,
							height: 650,
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
								text: "Amount",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Item Header Row Cell - Item No. (a)
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 0,
								y: 4110,
							},
							width: 300,
							height: 500,
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
								text: "Item NO. (a)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Item Header Row Cell - Federal Stock Number, Description, and Coding of Material and/or Services (b)
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 405,
								y: 4110,
							},
							width: 7200,
							height: 500,
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
								text: "Federal Stock Number, Description, and Coding of Material and/or Services (b)",
								allCaps: true,
								size: 12,
								bold: true,
								break: 1,
							}),
						],
					}),
					//Item Header Row Cell - Unit of Issue (c)
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 7705,
								y: 4110,
							},
							width: 450,
							height: 500,
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
								text: "Unit of Issue (c)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Item Header Row Cell - Quantity Requested (d)
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 8255,
								y: 4110,
							},
							width: 705,
							height: 500,
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
								text: "Quantity Requested (d)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Item Header Row Cell - Supply Action (e)
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9065,
								y: 4110,
							},
							width: 750,
							height: 500,
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
								text: "Supply Action (e)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Item Header Row Cell - Type Container (f)
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9920,
								y: 4110,
							},
							width: 800,
							height: 500,
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
								text: "Type Container (f)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Item Header Row Cell - Container Nos. (g)
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 10820,
								y: 4110,
							},
							width: 800,
							height: 500,
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
								text: "Container Nos. (g)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Item Header Row Cell - Unit Price (h)
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 11715,
								y: 4110,
							},
							width: 1000,
							height: 500,
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
								text: "Unit Price (h)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Item Header Row Cell - Total Cost (i)
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 12810,
								y: 4110,
							},
							width: 1000,
							height: 500,
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
								text: "Total Cost (i)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Item Data Cell - Numbered Item
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 0,
								y: 4590,
							},
							width: 300,
							height: 2550,
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
								text: "1",
								size: 14,
							}),
							new TextRun({
								text: "2",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "3",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "4",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "5",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "6",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "7",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "8",
								size: 14,
								break: 2,
							}),
						],
					}),
					//Item Data Cell - Federal Stock Number, Description, and Coding of Material and/or Services
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 405,
								y: 4590,
							},
							width: 7200,
							height: 2550,
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
								text: "736483 Test Equipment Batch 743",
								size: 14,
							}),
							new TextRun({
								text: "N/A",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "N/A",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "N/A",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "N/A",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "N/A",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "N/A",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "N/A",
								size: 14,
								break: 2,
							}),
						],
					}),
					//Item Data Cell - Unit of Issue
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 7705,
								y: 4590,
							},
							width: 450,
							height: 2550,
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
								text: "ABC",
								size: 14,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
						],
					}),
					//Item Data Cell - Quantity Requested
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 8255,
								y: 4590,
							},
							width: 705,
							height: 2550,
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
								text: "4",
								size: 14,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
						],
					}),
					//Item Data Cell - Supply Action
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9065,
								y: 4590,
							},
							width: 750,
							height: 2550,
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
								text: "XYZ",
								size: 14,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
						],
					}),
					//Item Data Cell - Type Container
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9920,
								y: 4590,
							},
							width: 800,
							height: 2550,
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
								text: "987A",
								size: 14,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
						],
					}),
					//Item Data Cell - Containers Nos.
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 10820,
								y: 4590,
							},
							width: 800,
							height: 2550,
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
								text: "7364",
								size: 14,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
						],
					}),
					//Item Data Cell - Unit Price
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 11715,
								y: 4590,
							},
							width: 1000,
							height: 2550,
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
								text: "$1,000.00",
								size: 14,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
						],
					}),
					//Item Data Cell - Total Cost
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 12810,
								y: 4590,
							},
							width: 1000,
							height: 2550,
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
								text: "$4,000.00",
								size: 14,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
							new TextRun({
								text: "---",
								size: 14,
								break: 2,
							}),
						],
					}),
					//Break Row Cell - #16 Transportation Via Mats or Msts Chargeable to:
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 7120,
							},
							width: 8155,
							height: 50,
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
								text: "#16 Transportation Via Mats or Msts Chargeable to: ",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "NLAF",
								size: 10,
								bold: true,
							}),
						],
					}),
					//Break Row Cell - #17 Special Handling
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8260,
								y: 7120,
							},
							width: 5550,
							height: 50,
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
								text: "#17 Special Handling: ",
								allCaps: true,
								size: 10,
							}),
							new TextRun({
								text: "N/A",
								size: 10,
								bold: true,
							}),
						],
					}),
					//Final Row Cell - #18 Recapitulation of Shipment
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 0,
								y: 7295,
							},
							width: 9305,
							height: 200,
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
								text: "#18 Recapitulation of Shipment",
								allCaps: true,
								size: 12,
							}),
						],
					}),
					//Final Row Cell - Issued By
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 7495,
							},
							width: 900,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								text: "Issued By:",
								allCaps: true,
								size: 12,
							}),
						],
					}),
					//Final Row Cell - Checked By
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 8225,
							},
							width: 900,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								text: "Checked By:",
								allCaps: true,
								size: 12,
							}),
						],
					}),
					//Final Row Cell - Packed
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 0,
								y: 8965,
							},
							width: 900,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								text: "Packed By:",
								allCaps: true,
								size: 12,
							}),
						],
					}),
					//Final Row Header Cell - Total Containers
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1010,
								y: 7495,
							},
							width: 850,
							height: 320,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "Total Containers",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Row Column Cells - Total Containers 1
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1010,
								y: 7805,
							},
							width: 850,
							height: 435,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Containers 2
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1010,
								y: 8225,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Containers 3
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1010,
								y: 8595,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Containers 4
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1010,
								y: 8965,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Containers 5
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1010,
								y: 9335,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Header Cell - Type Containers
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1960,
								y: 7495,
							},
							width: 850,
							height: 320,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "Type Containers",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Row Column Cells - Type Containers 1
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1960,
								y: 7805,
							},
							width: 850,
							height: 435,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Type Containers 2
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1960,
								y: 8225,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Type Containers 3
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1960,
								y: 8595,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Type Containers 4
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1960,
								y: 8965,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Type Containers 5
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 1960,
								y: 9335,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Header Cell - Description
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 2910,
								y: 7495,
							},
							width: 4500,
							height: 320,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "Description",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Row Column Cells - Description 1
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 2910,
								y: 7805,
							},
							width: 4500,
							height: 435,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Description 2
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 2910,
								y: 8225,
							},
							width: 4500,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Description 3
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 2910,
								y: 8595,
							},
							width: 4500,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Description 4
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 2910,
								y: 8965,
							},
							width: 4500,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Description 5
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 2910,
								y: 9335,
							},
							width: 4500,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Header Cell - Total Weight
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 7505,
								y: 7495,
							},
							width: 850,
							height: 320,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "Total Weight",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Row Column Cells - Total Weight 1
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 7505,
								y: 7805,
							},
							width: 850,
							height: 435,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Weight 2
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 7505,
								y: 8225,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Weight 3
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 7505,
								y: 8595,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Weight 4
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 7505,
								y: 8965,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Weight 5
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 7505,
								y: 9335,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Header Cell - Total Cube
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8455,
								y: 7495,
							},
							width: 850,
							height: 320,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "Total Cube",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Row Column Cells - Total Cube 1
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 8455,
								y: 7805,
							},
							width: 850,
							height: 435,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Cube 2
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 8455,
								y: 8225,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Cube 3
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8455,
								y: 8595,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Cube 4
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8455,
								y: 8965,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Column Cells - Total Cube 5
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 8455,
								y: 9335,
							},
							width: 850,
							height: 380,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "",
							}),
						],
					}),
					//Final Row Break - #19 Receipt
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9420,
								y: 7295,
							},
							width: 4390,
							height: 200,
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
								text: "#19 Receipt",
								allCaps: true,
								size: 12,
							}),
						],
					}),
					//Receipt Break
					new Paragraph({
						alignment: AlignmentType.CENTER,
						shading: {
							type: ShadingType.PERCENT_5,
							fill: "#DADADA",
						},
						frame: {
							position: {
								x: 9420,
								y: 7295,
							},
							width: 5,
							height: 2420,
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
								text: "",
								allCaps: true,
								size: 12,
							}),
						],
					}),
					//Final Section Row Alpha - Containers Received Except as Noted
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9540,
								y: 7495,
							},
							width: 900,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "Containers Received Except as Noted",
								allCaps: true,
								size: 13,
							}),
						],
					}),
					//Final Section Row Bravo - Quantities Received Except as Noted
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9540,
								y: 8225,
							},
							width: 900,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "Quantities Received Except as Noted",
								allCaps: true,
								size: 13,
							}),
						],
					}),
					//Final Section Row Charlie - Posted
					new Paragraph({
						alignment: AlignmentType.CENTER,
						frame: {
							position: {
								x: 9540,
								y: 8965,
							},
							width: 900,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "Posted",
								allCaps: true,
								size: 13,
								break: 1,
							}),
						],
					}),
					//Final Section Row Alpha - Date
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 10540,
								y: 7495,
							},
							width: 1000,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "Date (yymmdd)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Section Row Bravo - Date
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 10540,
								y: 8225,
							},
							width: 1000,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "Date (yymmdd)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Section Row Charlie - Date
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 10540,
								y: 8965,
							},
							width: 1000,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
							},
						},
						border: {
							top: {
								color: "auto",
								space: 1,
								value: "single",
								size: 0,
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
								size: 0,
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
								text: "Date (yymmdd)",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Section Row Alpha - By
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 11655,
								y: 7495,
							},
							width: 400,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "By:",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Section Row Bravo - By
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 11655,
								y: 8225,
							},
							width: 400,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "By:",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Section Row Charlie - By
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 11655,
								y: 8965,
							},
							width: 400,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "By:",
								allCaps: true,
								size: 10,
							}),
						],
					}),
					//Final Section Row Alpha - Sheet Total
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 12165,
								y: 7495,
							},
							width: 1645,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "Sheet Total",
								allCaps: true,
								size: 13,
							}),
						],
					}),
					//Final Section Row Bravo - Grand Total
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 12165,
								y: 8225,
							},
							width: 1645,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "Grand Total",
								allCaps: true,
								size: 13,
							}),
						],
					}),
					//Final Section Row Charlie - #20 Receivers Voucher No.
					new Paragraph({
						alignment: AlignmentType.LEFT,
						frame: {
							position: {
								x: 12165,
								y: 8965,
							},
							width: 1645,
							height: 750,
							anchor: {
								horizontal: FrameAnchorType.MARGIN,
								vertical: FrameAnchorType.MARGIN,
							},
							alignment: {
								x: HorizontalPositionAlign.CENTER,
								y: VerticalPositionAlign.CENTER,
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
								size: 0,
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
								text: "#20 Receivers Voucher No.",
								allCaps: true,
								size: 11,
							}),
						],
					}),
				],
			},
		],
	});

	Packer.toBlob(doc).then((blob) => {
		saveAs(blob, "Test Document");
	});
};

export default generateDoc;
