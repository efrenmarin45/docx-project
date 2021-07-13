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
                    //Disclosure Main Text
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
								y: 180,
							},
							width: 8010,
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
								y: 180,
							},
							width: 1000,
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
								x: 9220,
								y: 180,
							},
							width: 1000,
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
								x: 10325,
								y: 180,
							},
							width: 1200,
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
								x: 11630,
								y: 180,
							},
							width: 2180,
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
								y: 565,
							},
							width: 3415,
							height: 610,
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
								x: 11635,
								y: 565,
							},
							width: 2175,
							height: 610,
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
							width: 8010,
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
							height: 515,
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
							width: 2800,
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
								x: 11010,
								y: 1660,
							},
							width: 2800,
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

                    //SHIP TO Row Address

                    //SHIP TO Row Cell - #12 Date Shipped (YYMMDD)

                    //SHIP TO Row Cell - 11b EMPTY

                    //SHIP TO Row Cell - #13 Mode of Shipment

                    //SHIP TO Row Cell - #14 Bill of Landing Number

                    //SHIP TO Row Cell - #15 Air Movement Designator or Port Reference Number
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
