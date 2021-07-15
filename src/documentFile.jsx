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
								text: "FormApprovedOMBNo. 0704-0246",
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
								text: "Public reporting burden for this collection of information is estimated to average 1 hour per response, including the time for reviewing instructions, searching existing data sources, gathering and maintaining the data needed, and completing and reviewing the collection of information. Send comments regarding this burden estimate or any other aspect of this collection of information, including suggestions for reducing this burden, to Washington Headquarters Services, Directorate for Information Operations and Reports, 1215 Jefferson Davis Highway, Suite 1204, Arlington, VA 22202-4302, and to the Office of Management and Budget, Paperwork Reduction Project (0704-0246), Washington, DC20503.",
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
                            })
                        ]
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
                            })
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
								text: "TRIDENT REFIT FACILITY, CODE 532",
                                size: 18,
							}),
                            new TextRun({
                                text: "BLDG 4027",
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
                                text: "TyKit Program",
                                allCaps: true,
                                break: 1,
                                size: 18,
                            }),
                        ]
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
                        ]
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
                        ]
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
                        ]
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
								text: "DELPHINUS ENGINEERING",
                                size: 18,
							}),
                            new TextRun({
                                text: "3745 PROGRESS ROAD",
                                size: 18,
                                break: 1,
                            }),
                            new TextRun({
                                text: "NORFOLK, VA 23502",
                                size: 18,
                                break: 1,
                            }),
                            new TextRun({
                                text: "POC: RALPH TYLER (757) 588-8364 x360",
                                size: 18,
                                break: 1,
                            }),
                            new TextRun({
                                text: "** E-MAIL TRACKING INFO TO LEE.SAVELSON@NAVY.MIL",
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                                size: 14,
                            }),
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                        ]
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
                                size: 10,
                            }),
							new TextRun({
								text: "NLAF",
								size: 10,
								bold: true,
							})
                        ]
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
                                size: 10,
                            }),
							new TextRun({
								text: "N/A",
								size: 10,
								bold: true,
							})
                        ]
                    }),
					//Final Row Cell - #18 Recapitulation of Shipment

					//Final Row Cell - Issued By

					//Final Row Cell - Checked By

					//Final Row Cell - Packed

					//Final Row Header Cell - Total Containers

					//Final Row Column Cells - Total Containers

					//Final Row Header Cell - Type Containers

					//Final Row Column Cells - Type Containers

					//Final Row Header Cell - Description

					//Final Row Column Cells - Description

					//Final Row Header Cell - Total Weight

					//Final Row Column Cells - Total Weight

					//Final Row Header Cell - Total Cube

					//Final Row Column Cells - Total Cube

					//Final Row Break - #19 Receipt

					//Final Section Row Alpha - Containers Received Except as Noted

					//Final Section Row Alpha - Date

					//Final Section Row Alpha - By

					//Final Section Row Alpha - Sheet Total

					//Final Section Row Bravo - Quantities Received Except as Noted

					//Final Section Row Bravo - Date

					//Final Section Row Bravo - By

					//Final Section Row Bravo - Grand Total

					//Final Section Row Charlie - Posted

					//Final Section Row Charlie - Date

					//Final Section Row Charlie - By

					//Final Section Row Charlie - #20 Receivers Voucher No.
				],
			},
		],
	});

	Packer.toBlob(doc).then((blob) => {
		saveAs(blob, "Test Document");
	});
};

export default generateDoc;
