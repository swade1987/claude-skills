const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
    BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
    VerticalAlign, TableLayoutType } = require('docx');
const fs = require('fs');

// Elite Color Palette - Dark navy & gold for premium feel
const NAVY_DARK = "0D1B2A";
const NAVY_MID = "1B3A4B";
const NAVY_LIGHT = "2D5A6B";
const GOLD = "C9A227";
const GOLD_LIGHT = "F4E9CD";
const WHITE = "FFFFFF";
const LIGHT_GRAY = "F7F9FC";
const MID_GRAY = "E5E9F0";
const DARK_GRAY = "4A5568";
const TEXT_DARK = "1A202C";
const SUCCESS_GREEN = "059669";
const SUCCESS_LIGHT = "D1FAE5";
const WARNING_ORANGE = "D97706";
const WARNING_LIGHT = "FEF3C7";

// Border styles
const noBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const lightBorder = { style: BorderStyle.SINGLE, size: 1, color: MID_GRAY };
const lightBorders = { top: lightBorder, bottom: lightBorder, left: lightBorder, right: lightBorder };
const goldBorder = { style: BorderStyle.SINGLE, size: 2, color: GOLD };
const goldBorders = { top: goldBorder, bottom: goldBorder, left: goldBorder, right: goldBorder };
const navyBorder = { style: BorderStyle.SINGLE, size: 2, color: NAVY_DARK };
const navyBorders = { top: navyBorder, bottom: navyBorder, left: navyBorder, right: navyBorder };
const successBorder = { style: BorderStyle.SINGLE, size: 2, color: SUCCESS_GREEN };
const successBorders = { top: successBorder, bottom: successBorder, left: successBorder, right: successBorder };

// Page dimensions (A4)
const PAGE_WIDTH = 11906;
const PAGE_HEIGHT = 16838;
const MARGIN = 1200;
const CONTENT_WIDTH = PAGE_WIDTH - (MARGIN * 2); // 9506 DXA

// Helpers
const spacer = (height = 200) => new Paragraph({ spacing: { after: height }, children: [] });

const createParagraph = (text, options = {}) => {
    return new Paragraph({
        spacing: { after: 160, line: 276 },
        ...options,
        children: [new TextRun({
            text,
            font: "Calibri",
            size: 22,
            color: TEXT_DARK,
            ...options.textOptions
        })]
    });
};

const createMultiParagraph = (runs, options = {}) => {
    return new Paragraph({
        spacing: { after: 160, line: 276 },
        ...options,
        children: runs.map(run => new TextRun({
            font: "Calibri",
            size: 22,
            color: TEXT_DARK,
            ...run
        }))
    });
};

const createHeading = (text, level = 1) => {
    const sizes = { 1: 32, 2: 26, 3: 22 };
    const colors = { 1: NAVY_DARK, 2: NAVY_MID, 3: DARK_GRAY };
    const spacingBefore = { 1: 400, 2: 300, 3: 200 };
    const spacingAfter = { 1: 200, 2: 160, 3: 120 };

    return new Paragraph({
        spacing: { before: spacingBefore[level], after: spacingAfter[level] },
        children: [new TextRun({
            text,
            bold: true,
            font: "Calibri Light",
            size: sizes[level],
            color: colors[level]
        })]
    });
};

const createBullet = (text, reference = "bullets") => {
    return new Paragraph({
        numbering: { reference, level: 0 },
        spacing: { after: 100, line: 276 },
        children: [new TextRun({ text, font: "Calibri", size: 22, color: TEXT_DARK })]
    });
};

const createNumberedItem = (text, reference = "numbers") => {
    return new Paragraph({
        numbering: { reference, level: 0 },
        spacing: { after: 100, line: 276 },
        children: [new TextRun({ text, font: "Calibri", size: 22, color: TEXT_DARK })]
    });
};

// Elite table styles
const createEliteTable = (headers, rows, options = {}) => {
    const colCount = headers.length;
    const colWidth = Math.floor(CONTENT_WIDTH / colCount);
    const columnWidths = options.columnWidths || Array(colCount).fill(colWidth);

    const headerRow = new TableRow({
        tableHeader: true,
        children: headers.map((header, i) => new TableCell({
            borders: noBorders,
            width: { size: columnWidths[i], type: WidthType.DXA },
            shading: { fill: NAVY_DARK, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 120, right: 120 },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
                alignment: options.headerAlign || AlignmentType.LEFT,
                children: [new TextRun({ text: header, bold: true, font: "Calibri", size: 20, color: WHITE })]
            })]
        }))
    });

    const dataRows = rows.map((row, rowIndex) => new TableRow({
        children: row.map((cell, i) => {
            const isHighlighted = options.highlightRow === rowIndex;
            const cellText = typeof cell === 'object' ? cell.text : cell;
            const cellBold = typeof cell === 'object' ? cell.bold : false;
            const cellColor = typeof cell === 'object' && cell.color ? cell.color : TEXT_DARK;

            return new TableCell({
                borders: {
                    top: noBorder,
                    bottom: { style: BorderStyle.SINGLE, size: 1, color: MID_GRAY },
                    left: noBorder,
                    right: noBorder
                },
                width: { size: columnWidths[i], type: WidthType.DXA },
                shading: { fill: isHighlighted ? GOLD_LIGHT : (rowIndex % 2 === 0 ? WHITE : LIGHT_GRAY), type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({
                    alignment: options.cellAlign && options.cellAlign[i] ? options.cellAlign[i] : AlignmentType.LEFT,
                    children: [new TextRun({
                        text: cellText,
                        bold: cellBold || isHighlighted,
                        font: "Calibri",
                        size: 20,
                        color: cellColor
                    })]
                })]
            });
        })
    }));

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths,
        rows: [headerRow, ...dataRows]
    });
};

// Callout box
const createCalloutBox = (title, content, bgColor, borderColor, titleBgColor) => {
    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [CONTENT_WIDTH],
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 2, color: borderColor },
                            left: { style: BorderStyle.SINGLE, size: 2, color: borderColor },
                            right: { style: BorderStyle.SINGLE, size: 2, color: borderColor },
                            bottom: noBorder
                        },
                        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
                        shading: { fill: titleBgColor || borderColor, type: ShadingType.CLEAR },
                        margins: { top: 100, bottom: 100, left: 160, right: 160 },
                        children: [new Paragraph({
                            children: [new TextRun({ text: title, bold: true, font: "Calibri", size: 24, color: WHITE })]
                        })]
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: noBorder,
                            left: { style: BorderStyle.SINGLE, size: 2, color: borderColor },
                            right: { style: BorderStyle.SINGLE, size: 2, color: borderColor },
                            bottom: { style: BorderStyle.SINGLE, size: 2, color: borderColor }
                        },
                        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
                        shading: { fill: bgColor, type: ShadingType.CLEAR },
                        margins: { top: 140, bottom: 140, left: 160, right: 160 },
                        children: content
                    })
                ]
            })
        ]
    });
};

// Option pricing box
const createOptionBox = (number, title, price, duration, isRecommended = false) => {
    const borderColor = isRecommended ? GOLD : NAVY_DARK;
    const headerBg = isRecommended ? GOLD : NAVY_DARK;
    const badge = isRecommended ? "  ★ RECOMMENDED" : "";

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [CONTENT_WIDTH],
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 3, color: borderColor },
                            left: { style: BorderStyle.SINGLE, size: 3, color: borderColor },
                            right: { style: BorderStyle.SINGLE, size: 3, color: borderColor },
                            bottom: noBorder
                        },
                        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
                        shading: { fill: headerBg, type: ShadingType.CLEAR },
                        margins: { top: 120, bottom: 120, left: 200, right: 200 },
                        children: [new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                                new TextRun({ text: `OPTION ${number}: `, font: "Calibri", size: 24, color: WHITE }),
                                new TextRun({ text: title, bold: true, font: "Calibri", size: 28, color: WHITE }),
                                new TextRun({ text: badge, bold: true, font: "Calibri", size: 20, color: WHITE })
                            ]
                        })]
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: noBorder,
                            left: { style: BorderStyle.SINGLE, size: 3, color: borderColor },
                            right: { style: BorderStyle.SINGLE, size: 3, color: borderColor },
                            bottom: { style: BorderStyle.SINGLE, size: 3, color: borderColor }
                        },
                        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
                        shading: { fill: isRecommended ? GOLD_LIGHT : LIGHT_GRAY, type: ShadingType.CLEAR },
                        margins: { top: 140, bottom: 140, left: 200, right: 200 },
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 60 },
                                children: [new TextRun({ text: price, bold: true, font: "Calibri Light", size: 48, color: isRecommended ? NAVY_DARK : NAVY_DARK })]
                            }),
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                children: [new TextRun({ text: duration, font: "Calibri", size: 22, color: DARK_GRAY })]
                            })
                        ]
                    })
                ]
            })
        ]
    });
};

// Stats box
const createStatsBox = (stats) => {
    const colWidth = Math.floor(CONTENT_WIDTH / stats.length);
    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: Array(stats.length).fill(colWidth),
        rows: [
            new TableRow({
                children: stats.map((stat, i) => new TableCell({
                    borders: {
                        top: { style: BorderStyle.SINGLE, size: 2, color: NAVY_DARK },
                        bottom: { style: BorderStyle.SINGLE, size: 2, color: NAVY_DARK },
                        left: i === 0 ? { style: BorderStyle.SINGLE, size: 2, color: NAVY_DARK } : noBorder,
                        right: i === stats.length - 1 ? { style: BorderStyle.SINGLE, size: 2, color: NAVY_DARK } : noBorder
                    },
                    width: { size: colWidth, type: WidthType.DXA },
                    shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
                    margins: { top: 160, bottom: 160, left: 120, right: 120 },
                    verticalAlign: VerticalAlign.CENTER,
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 40 },
                            children: [new TextRun({ text: stat.value, bold: true, font: "Calibri Light", size: 52, color: NAVY_DARK })]
                        }),
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ text: stat.label, font: "Calibri", size: 18, color: DARK_GRAY })]
                        })
                    ]
                }))
            })
        ]
    });
};

// Value stack table
const createValueStack = (items, totalValue, yourPrice) => {
    const rows = items.map(item => [item.component, { text: item.value, bold: false }]);
    rows.push([{ text: "TOTAL VALUE", bold: true }, { text: totalValue, bold: true, color: NAVY_DARK }]);
    rows.push([{ text: "YOUR INVESTMENT", bold: true }, { text: yourPrice, bold: true, color: SUCCESS_GREEN }]);

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [Math.floor(CONTENT_WIDTH * 0.7), Math.floor(CONTENT_WIDTH * 0.3)],
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: noBorders,
                        width: { size: Math.floor(CONTENT_WIDTH * 0.7), type: WidthType.DXA },
                        shading: { fill: NAVY_DARK, type: ShadingType.CLEAR },
                        margins: { top: 80, bottom: 80, left: 120, right: 120 },
                        children: [new Paragraph({ children: [new TextRun({ text: "Component", bold: true, font: "Calibri", size: 20, color: WHITE })] })]
                    }),
                    new TableCell({
                        borders: noBorders,
                        width: { size: Math.floor(CONTENT_WIDTH * 0.3), type: WidthType.DXA },
                        shading: { fill: NAVY_DARK, type: ShadingType.CLEAR },
                        margins: { top: 80, bottom: 80, left: 120, right: 120 },
                        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Value", bold: true, font: "Calibri", size: 20, color: WHITE })] })]
                    })
                ]
            }),
            ...rows.map((row, i) => {
                const isTotal = i === rows.length - 2;
                const isPrice = i === rows.length - 1;
                const bgColor = isTotal || isPrice ? (isPrice ? SUCCESS_LIGHT : GOLD_LIGHT) : (i % 2 === 0 ? WHITE : LIGHT_GRAY);

                return new TableRow({
                    children: [
                        new TableCell({
                            borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 1, color: MID_GRAY }, left: noBorder, right: noBorder },
                            width: { size: Math.floor(CONTENT_WIDTH * 0.7), type: WidthType.DXA },
                            shading: { fill: bgColor, type: ShadingType.CLEAR },
                            margins: { top: 70, bottom: 70, left: 120, right: 120 },
                            children: [new Paragraph({ children: [new TextRun({ text: typeof row[0] === 'object' ? row[0].text : row[0], bold: typeof row[0] === 'object' ? row[0].bold : false, font: "Calibri", size: 20, color: TEXT_DARK })] })]
                        }),
                        new TableCell({
                            borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 1, color: MID_GRAY }, left: noBorder, right: noBorder },
                            width: { size: Math.floor(CONTENT_WIDTH * 0.3), type: WidthType.DXA },
                            shading: { fill: bgColor, type: ShadingType.CLEAR },
                            margins: { top: 70, bottom: 70, left: 120, right: 120 },
                            children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: typeof row[1] === 'object' ? row[1].text : row[1], bold: typeof row[1] === 'object' ? row[1].bold : false, font: "Calibri", size: 20, color: typeof row[1] === 'object' && row[1].color ? row[1].color : TEXT_DARK })] })]
                        })
                    ]
                });
            })
        ]
    });
};

// Main document
const doc = new Document({
    styles: {
        default: {
            document: {
                run: { font: "Calibri", size: 22 }
            }
        },
        paragraphStyles: [
            {
                id: "Heading1",
                name: "Heading 1",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: { size: 32, bold: true, font: "Calibri Light", color: NAVY_DARK },
                paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 }
            },
            {
                id: "Heading2",
                name: "Heading 2",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: { size: 26, bold: true, font: "Calibri Light", color: NAVY_MID },
                paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 1 }
            },
            {
                id: "Heading3",
                name: "Heading 3",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: { size: 22, bold: true, font: "Calibri", color: DARK_GRAY },
                paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 }
            }
        ]
    },
    numbering: {
        config: [
            {
                reference: "bullets",
                levels: [{
                    level: 0,
                    format: LevelFormat.BULLET,
                    text: "●",
                    alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
            {
                reference: "numbers",
                levels: [{
                    level: 0,
                    format: LevelFormat.DECIMAL,
                    text: "%1.",
                    alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
            {
                reference: "checks",
                levels: [{
                    level: 0,
                    format: LevelFormat.BULLET,
                    text: "✓",
                    alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            }
        ]
    },
    sections: [{
        properties: {
            page: {
                size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
                margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
            }
        },
        headers: {
            default: new Header({
                children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun({ text: "PLATFORM FIX", font: "Calibri Light", size: 18, color: NAVY_DARK }),
                        new TextRun({ text: "  |  ", font: "Calibri", size: 18, color: MID_GRAY }),
                        new TextRun({ text: "Confidential Proposal", font: "Calibri", size: 18, color: DARK_GRAY })
                    ]
                })]
            })
        },
        footers: {
            default: new Footer({
                children: [
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        columnWidths: [Math.floor(CONTENT_WIDTH / 2), Math.floor(CONTENT_WIDTH / 2)],
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders: noBorders,
                                        width: { size: Math.floor(CONTENT_WIDTH / 2), type: WidthType.DXA },
                                        children: [new Paragraph({
                                            children: [new TextRun({ text: "steve@platformfix.com", font: "Calibri", size: 16, color: DARK_GRAY })]
                                        })]
                                    }),
                                    new TableCell({
                                        borders: noBorders,
                                        width: { size: Math.floor(CONTENT_WIDTH / 2), type: WidthType.DXA },
                                        children: [new Paragraph({
                                            alignment: AlignmentType.RIGHT,
                                            children: [
                                                new TextRun({ text: "Page ", font: "Calibri", size: 16, color: DARK_GRAY }),
                                                new TextRun({ children: [PageNumber.CURRENT], font: "Calibri", size: 16, color: DARK_GRAY }),
                                                new TextRun({ text: " of ", font: "Calibri", size: 16, color: DARK_GRAY }),
                                                new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Calibri", size: 16, color: DARK_GRAY })
                                            ]
                                        })]
                                    })
                                ]
                            })
                        ]
                    })
                ]
            })
        },
        children: [
            // ========== COVER PAGE ==========
            spacer(600),

            // Logo/Brand
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 },
                children: [new TextRun({ text: "PLATFORM FIX", bold: true, font: "Calibri Light", size: 56, color: NAVY_DARK })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 600 },
                children: [new TextRun({ text: "Platform Complexity Transformation", font: "Calibri Light", size: 28, color: DARK_GRAY })]
            }),

            // Divider line
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                columnWidths: [CONTENT_WIDTH],
                rows: [new TableRow({
                    children: [new TableCell({
                        borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 3, color: GOLD }, left: noBorder, right: noBorder },
                        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
                        children: [new Paragraph({ children: [] })]
                    })]
                })]
            }),

            spacer(600),

            // Client box
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                columnWidths: [CONTENT_WIDTH],
                rows: [new TableRow({
                    children: [new TableCell({
                        borders: navyBorders,
                        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
                        shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
                        margins: { top: 300, bottom: 300, left: 400, right: 400 },
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 80 },
                                children: [new TextRun({ text: "PROPOSAL", font: "Calibri", size: 20, color: DARK_GRAY, allCaps: true })]
                            }),
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 80 },
                                children: [new TextRun({ text: "Prepared Exclusively For", font: "Calibri", size: 20, color: DARK_GRAY })]
                            }),
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 80 },
                                children: [new TextRun({ text: "WEISSHORN C5I / CYD AG", bold: true, font: "Calibri Light", size: 36, color: NAVY_DARK })]
                            }),
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                children: [new TextRun({ text: "January 2026", font: "Calibri", size: 22, color: DARK_GRAY })]
                            })
                        ]
                    })]
                })]
            }),

            spacer(800),

            // Prepared by
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 },
                children: [new TextRun({ text: "Prepared by", font: "Calibri", size: 18, color: DARK_GRAY })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [new TextRun({ text: "Steve Wade", bold: true, font: "Calibri", size: 24, color: NAVY_DARK })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [new TextRun({ text: "Founder, Platform Fix", font: "Calibri", size: 20, color: DARK_GRAY })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "steve@platformfix.com", font: "Calibri", size: 20, color: NAVY_MID })]
            }),

            // ========== PAGE 2: EXECUTIVE SUMMARY ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Executive Summary", 1),
            createParagraph("Following our discovery call on January 27th, 2026, Platform Fix is pleased to present this proposal for transforming your sovereign on-premise Kubernetes platform."),
            createParagraph("Weisshorn C5I operates a critical platform for Swiss cyber defence, currently supporting 3 engineering teams with plans to scale to 30 teams by end of 2026. The platform has been running successfully for 2+ years, but faces scaling challenges around team onboarding, knowledge concentration, and deployment complexity."),
            createParagraph("This document presents three engagement options designed to address your specific challenges, from diagnostic assessment through full transformation with ongoing support."),

            spacer(300),

            // Client info table
            createHeading("Engagement Details", 2),
            createEliteTable(
                ["Field", "Details"],
                [
                    ["Client", "Weisshorn C5I / Cyd AG"],
                    ["Location", "Switzerland"],
                    ["Primary Contact", "Benoit Perroud"],
                    ["Platform Fix Contact", "Steve Wade"],
                    ["Document Date", "January 28, 2026"],
                    ["Valid Until", "February 28, 2026"]
                ],
                { columnWidths: [3500, CONTENT_WIDTH - 3500] }
            ),

            // ========== PAGE 3: BUSINESS CONTEXT ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Business Context", 1),

            createHeading("Current Situation", 2),
            createBullet("Sovereign on-premise Kubernetes platform for Swiss cyber defence applications"),
            createBullet("Platform operational for 2+ years, started with Kubernetes in 2017"),
            createBullet("Team of 8 people (3 FTEs dedicated to platform, 5 on applications)"),
            createBullet("Currently supporting 3 engineering teams"),
            createBullet("Goal: Scale to 30 teams by end of 2026"),

            spacer(200),

            createHeading("Critical Challenges Identified", 2),
            createBullet("Knowledge concentration: Platform expertise concentrated in single individual (bottleneck risk)"),
            createBullet("Onboarding friction: New teams require extensive hand-holding, slowing adoption"),
            createBullet("Deployment complexity: 20-40% productivity loss in deployment phase"),
            createBullet("Testing gaps: Testing primarily happens in production"),
            createBullet("Manual processes: Configuration tweaking required between code push and production"),

            spacer(200),

            createHeading("Desired Outcomes", 2),
            createBullet("Platform that works without heroics. Teams ship features, not fight fires"),
            createBullet("Self-service onboarding. New teams productive in days, not weeks"),
            createBullet("Reduced dependency. Platform runs without a single point of failure"),
            createBullet("Streamlined deployments. Push to main goes straight to production"),
            createBullet("Scalable foundation. Ready to support 30 teams by end of 2026"),

            // ========== PAGE 4: COST OF WAITING ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("The Cost of Waiting", 1),

            createCalloutBox(
                "EVERY MONTH WITHOUT ACTION COSTS YOU",
                [
                    new Paragraph({
                        spacing: { after: 160 },
                        children: [new TextRun({ text: "Based on our discovery call, here's what the current situation costs:", font: "Calibri", size: 22, color: TEXT_DARK })]
                    }),
                    createEliteTable(
                        ["Cost Factor", "Monthly/Annual Impact"],
                        [
                            ["Platform team productivity loss (20-40%)", "CHF 7,500 - 15,000/month"],
                            ["Delayed feature delivery to customers", "Unquantified but significant"],
                            ["Risk of shadow IT if teams frustrated", "Security + compliance risk"],
                            ["Onboarding bottleneck", "27 teams × weeks of delay"],
                            [{ text: "CONSERVATIVE ANNUAL COST", bold: true }, { text: "CHF 90,000 - 180,000+", bold: true, color: WARNING_ORANGE }]
                        ],
                        { columnWidths: [Math.floor(CONTENT_WIDTH * 0.6), Math.floor(CONTENT_WIDTH * 0.4)], highlightRow: 4 }
                    )
                ],
                WARNING_LIGHT,
                WARNING_ORANGE,
                WARNING_ORANGE
            ),

            spacer(300),

            createMultiParagraph([
                { text: "Every month you delay, you're spending another CHF 7,500-15,000 on productivity drag alone. By the time you've manually onboarded your 10th team without fixing this, you'll have burned through approximately CHF 180,000. " },
                { text: "six times the cost of Foundation.", bold: true }
            ]),

            createMultiParagraph([
                { text: "The Foundation engagement (£35,000 ≈ CHF 38,000) " },
                { text: "pays for itself within 3-5 months.", bold: true }
            ]),

            // ========== PAGE 5: WHY PLATFORM FIX ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Why Platform Fix", 1),

            createHeading("We Delete, We Don't Add", 2),
            createParagraph("Most consultancies come in, add more tools, more abstraction layers, more complexity, then leave you with a bigger mess than before. We do the opposite."),
            createMultiParagraph([
                { text: "Our core philosophy: " },
                { text: "Boring is Beautiful.", bold: true },
                { text: " We don't chase shiny Kubernetes operators. We ruthlessly eliminate everything that doesn't need to exist. The result? Platforms that actually work." }
            ]),

            spacer(200),

            createHeading("Engineers, Not MBAs", 2),
            createParagraph("Platform Fix is engineer-led. We've sat in your seat, managing platforms, fighting fires, onboarding teams. We're not business consultants theorising about Kubernetes. We've built, broken, and fixed platforms for over a decade."),

            spacer(200),

            createHeading("Track Record", 2),
            createStatsBox([
                { value: "53", label: "Platform Transformations" },
                { value: "40-70%", label: "Complexity Reduction" }
            ]),

            spacer(400),

            // Case Study
            createCalloutBox(
                "CASE STUDY: European FinTech Platform Transformation",
                [
                    createMultiParagraph([
                        { text: "Challenge: ", bold: true },
                        { text: "Series B FinTech with over-engineered Kubernetes platform. Platform team spending 60% of time on maintenance. New feature deployments taking 3+ days." }
                    ]),
                    createMultiParagraph([
                        { text: "Approach: ", bold: true },
                        { text: "6-week Foundation engagement. Deleted 12 unnecessary operators, consolidated 4 monitoring tools into 1, simplified deployment pipeline from 47 steps to 8." }
                    ]),
                    new Paragraph({
                        spacing: { after: 80 },
                        children: [new TextRun({ text: "Results:", bold: true, font: "Calibri", size: 22, color: TEXT_DARK })]
                    }),
                    new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 60 }, children: [new TextRun({ text: "Platform complexity reduced by 58%", font: "Calibri", size: 22, color: TEXT_DARK })] }),
                    new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 60 }, children: [new TextRun({ text: "Infrastructure costs reduced from £140k to £72k annually", font: "Calibri", size: 22, color: TEXT_DARK })] }),
                    new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 60 }, children: [new TextRun({ text: "Deployment time reduced from 3 days to 3 hours", font: "Calibri", size: 22, color: TEXT_DARK })] }),
                    new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 120 }, children: [new TextRun({ text: "Platform team reclaimed 40% of time for feature work", font: "Calibri", size: 22, color: TEXT_DARK })] }),
                    new Paragraph({
                        spacing: { before: 100 },
                        children: [
                            new TextRun({ text: "\"We thought we needed to add more tooling. Platform Fix showed us we needed to delete 70% of what we had.\"", italics: true, font: "Calibri", size: 22, color: TEXT_DARK }),
                            new TextRun({ text: ". VP Engineering", font: "Calibri", size: 22, color: DARK_GRAY })
                        ]
                    })
                ],
                LIGHT_GRAY,
                NAVY_DARK,
                NAVY_DARK
            ),

            // ========== PAGE 6: WHY THIS OFFER ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Why This Offer?", 1),
            createParagraph("You might be wondering why we're offering this level of transformation at this price."),

            spacer(100),

            createHeading("Two Reasons:", 2),

            createMultiParagraph([
                { text: "1. We're expanding in the Swiss market. ", bold: true },
                { text: "Weisshorn is exactly the type of client we want: sophisticated platform, real challenges, ambitious scale-up. Your success becomes our flagship case study for the region." }
            ]),
            spacer(100),
            createMultiParagraph([
                { text: "2. Your timeline aligns with ours. ", bold: true },
                { text: "We have availability in late February / early March. That's rare; we typically book 6-8 weeks out." }
            ]),

            spacer(200),

            createParagraph("This isn't discounted work. It's full-price work at a moment when the fit is perfect."),

            // ========== PAGE 7: OPTION 1 ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Engagement Options", 1),
            createParagraph("Based on our discovery conversation, we recommend one of three engagement levels. Each builds upon the previous, allowing you to choose the depth of transformation and support that matches your needs."),

            spacer(300),

            createOptionBox("1", "Platform Audit", "£7,500", "1 Week", false),

            spacer(200),

            new Paragraph({
                spacing: { after: 160 },
                children: [new TextRun({ text: "Diagnosis only. Understand exactly what's broken and what to fix first.", italics: true, font: "Calibri", size: 22, color: DARK_GRAY })]
            }),

            createHeading("What's Included:", 3),
            createBullet("Complexity Score™ comprehensive assessment"),
            createBullet("Cost analysis (annual waste quantified in CHF)"),
            createBullet("Quick wins document (3-5 immediate deletions, CHF 50k+ value)"),
            createBullet("12-month simplification roadmap"),
            createBullet("60-minute results presentation"),

            spacer(100),

            createHeading("What's NOT Included:", 3),
            createBullet("Implementation of recommendations"),
            createBullet("Team training"),
            createBullet("Ongoing support"),

            spacer(100),

            createHeading("Delivery:", 3),
            createBullet("Remote documentation review and calls"),
            createBullet("1-2 days on-site for team interviews (subject to access approval)"),

            spacer(100),

            createMultiParagraph([
                { text: "Outcome: ", bold: true },
                { text: "Complete clarity on what's broken and what to fix. Implementation is on you." }
            ]),

            spacer(100),

            createMultiParagraph([
                { text: "Upgrade Credit: ", bold: true },
                { text: "If you proceed to Foundation within 60 days, your Audit fee applies as credit toward the full engagement." }
            ]),

            // ========== PAGE 8: OPTION 2 ==========
            new Paragraph({ children: [new PageBreak()] }),

            createOptionBox("2", "Foundation", "£35,000", "6 Weeks", true),

            spacer(200),

            new Paragraph({
                spacing: { after: 160 },
                children: [new TextRun({ text: "Full transformation. We diagnose AND fix the problems, then train your team.", italics: true, font: "Calibri", size: 22, color: DARK_GRAY })]
            }),

            createHeading("What You Get: The Value Stack", 2),

            createValueStack(
                [
                    { component: "Complexity Score™ Comprehensive Audit", value: "£7,500" },
                    { component: "Cost Analysis & Waste Quantification", value: "£2,000" },
                    { component: "12-Month Simplification Roadmap", value: "£3,000" },
                    { component: "Full Platform Transformation Execution", value: "£15,000" },
                    { component: "Team Training & Knowledge Transfer", value: "£5,000" },
                    { component: "Platform Fix OS™ Implementation", value: "£4,000" },
                    { component: "Onboarding Playbook for New Teams", value: "£3,000" },
                    { component: "Complete Documentation & Runbooks", value: "£3,000" },
                    { component: "On-Time Delivery Guarantee", value: "£2,000" }
                ],
                "£44,500",
                "£35,000"
            ),

            spacer(200),

            createMultiParagraph([
                { text: "You're getting £44,500 of value for £35,000, a ", bold: true },
                { text: "29% saving.", bold: true, color: SUCCESS_GREEN }
            ]),

            // ========== PAGE 9: OPTION 2 CONTINUED ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Foundation: Delivery Timeline", 2),
            createBullet("Week 1-2: Discovery and audit (on-site in Switzerland)"),
            createBullet("Week 2-3: Quick wins implementation (CHF 50k+ value delivered)"),
            createBullet("Week 3-5: Core deletions (major complexity reduction)"),
            createBullet("Week 5-6: Documentation, training, handoff"),

            spacer(300),

            createHeading("Bonuses. Included With Foundation", 2),

            createEliteTable(
                ["Bonus", "Value", "Description"],
                [
                    ["90-Day Complexity Health Check", "£2,500", "Three months after completion, we reassess your Complexity Score™ to ensure improvements stick."],
                    ["Team Onboarding Template Pack", "£1,500", "Ready-to-use templates for onboarding your next 10 teams."],
                    [{ text: "TOTAL BONUS VALUE", bold: true }, { text: "£4,000", bold: true }, { text: "Yours free with Foundation", bold: true }]
                ],
                { columnWidths: [2800, 1200, CONTENT_WIDTH - 4000], highlightRow: 2 }
            ),

            spacer(300),

            createMultiParagraph([
                { text: "Outcome: ", bold: true },
                { text: "Platform complexity eliminated. Your team can onboard new teams without you. You take a holiday without your phone ringing." }
            ]),

            // ========== PAGE 10: GUARANTEE ==========
            new Paragraph({ children: [new PageBreak()] }),

            createCalloutBox(
                "THE PLATFORM FIX GUARANTEE",
                [
                    new Paragraph({
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "We deliver results, not \"best effort.\" No 200-page strategy decks. Actual measurable outcomes.", bold: true, font: "Calibri", size: 22, color: TEXT_DARK })]
                    }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        columnWidths: [CONTENT_WIDTH - 400],
                        rows: [
                            new TableRow({
                                children: [new TableCell({
                                    borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 1, color: MID_GRAY }, left: noBorder, right: noBorder },
                                    width: { size: CONTENT_WIDTH - 400, type: WidthType.DXA },
                                    margins: { top: 100, bottom: 100, left: 0, right: 0 },
                                    children: [
                                        new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "GUARANTEE 1: Waste Identification", bold: true, font: "Calibri", size: 22, color: SUCCESS_GREEN })] }),
                                        new Paragraph({ children: [new TextRun({ text: "By Week 2, we identify at least CHF 100,000 in annual cost reduction opportunities, or we refund £7,500 of your investment. You keep all the audit findings.", font: "Calibri", size: 22, color: TEXT_DARK })] })
                                    ]
                                })]
                            }),
                            new TableRow({
                                children: [new TableCell({
                                    borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 1, color: MID_GRAY }, left: noBorder, right: noBorder },
                                    width: { size: CONTENT_WIDTH - 400, type: WidthType.DXA },
                                    margins: { top: 100, bottom: 100, left: 0, right: 0 },
                                    children: [
                                        new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "GUARANTEE 2: On-Time Delivery", bold: true, font: "Calibri", size: 22, color: SUCCESS_GREEN })] }),
                                        new Paragraph({ children: [new TextRun({ text: "Complete in 6 weeks, or we continue working at no additional cost until done.", font: "Calibri", size: 22, color: TEXT_DARK })] })
                                    ]
                                })]
                            }),
                            new TableRow({
                                children: [new TableCell({
                                    borders: noBorders,
                                    width: { size: CONTENT_WIDTH - 400, type: WidthType.DXA },
                                    margins: { top: 100, bottom: 100, left: 0, right: 0 },
                                    children: [
                                        new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "GUARANTEE 3: Measurable Results", bold: true, font: "Calibri", size: 22, color: SUCCESS_GREEN })] }),
                                        new Paragraph({ children: [new TextRun({ text: "40-70% complexity reduction, measured by Complexity Score™ before and after. If we don't hit 40%, we continue working at no charge until we do.", font: "Calibri", size: 22, color: TEXT_DARK })] })
                                    ]
                                })]
                            })
                        ]
                    }),
                    new Paragraph({
                        spacing: { before: 160 },
                        children: [new TextRun({ text: "The risk is entirely on us. We've done this 53 times. We know what we're walking into.", italics: true, font: "Calibri", size: 22, color: DARK_GRAY })]
                    })
                ],
                SUCCESS_LIGHT,
                SUCCESS_GREEN,
                SUCCESS_GREEN
            ),

            // ========== PAGE 11: OPTION 3 ==========
            new Paragraph({ children: [new PageBreak()] }),

            createOptionBox("3", "Momentum", "£59,000", "6 Weeks + 6 Months Advisory", false),

            spacer(200),

            new Paragraph({
                spacing: { after: 160 },
                children: [new TextRun({ text: "Transformation plus ongoing partnership. We stay alongside you as you scale to 30 teams.", italics: true, font: "Calibri", size: 22, color: DARK_GRAY })]
            }),

            createHeading("What's Included:", 3),
            createBullet("Everything in Foundation (£35,000 value) including all bonuses"),
            createBullet("6 months Platform Excellence advisory (£4,000/month × 6)"),
            createBullet("Weekly 2-hour office hours"),
            createBullet("Monthly architecture reviews"),
            createBullet("Priority Slack support (24-hour response)"),
            createBullet("Quarterly best practices workshops"),
            createBullet("Continuous optimisation recommendations"),

            spacer(200),

            createHeading("Additional Momentum Bonus", 3),
            createEliteTable(
                ["Bonus", "Value", "Description"],
                [
                    ["Second 90-Day Health Check", "£2,500", "A second Complexity Score™ assessment at the 9-month mark."]
                ],
                { columnWidths: [2800, 1200, CONTENT_WIDTH - 4000] }
            ),

            spacer(200),

            createHeading("Why This Option:", 3),
            createParagraph("You're scaling from 3 to 30 teams by end of 2026. That's 27 team onboardings in 12 months. Momentum ensures we're alongside you for the critical first wave, preventing complexity creep and course-correcting as you scale."),

            spacer(100),

            createMultiParagraph([
                { text: "Outcome: ", bold: true },
                { text: "Platform transformed AND ongoing expert partnership as you 10x your platform's user base." }
            ]),

            // ========== PAGE 12: COMPARISON ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Option Comparison", 1),

            createEliteTable(
                ["", "Audit", "Foundation ★", "Momentum"],
                [
                    ["Investment", "£7,500", { text: "£35,000", bold: true }, "£59,000"],
                    ["Total Value", "£7,500", { text: "£48,500", bold: true }, "£67,000"],
                    ["Duration", "1 week", "6 weeks", "6 weeks + 6 months"],
                    ["Complexity Score™ Audit", "✓", "✓", "✓"],
                    ["Implementation", ". ", "✓", "✓"],
                    ["Team Training", ". ", "✓", "✓"],
                    ["Results Guarantee", ". ", "✓", "✓"],
                    ["90-Day Health Check", ". ", "✓", "✓"],
                    ["Onboarding Templates", ". ", "✓", "✓"],
                    ["Ongoing Support", ". ", ". ", "6 months"],
                    ["On-Site Presence", "1-2 days", "2 weeks", "2 weeks"],
                    ["Payback Period", "N/A", { text: "3-5 months", bold: true, color: SUCCESS_GREEN }, "5-7 months"]
                ],
                { columnWidths: [2800, 2000, 2400, 2306], highlightRow: null }
            ),

            spacer(400),

            createHeading("Which Option Is Right For You?", 2),

            createEliteTable(
                ["If you want to...", "Choose"],
                [
                    ["Validate the approach before committing", "Audit (£7.5k)"],
                    [{ text: "Solve the problem and train your team", bold: true }, { text: "Foundation (£35k) ★", bold: true }],
                    ["Transform + ongoing support through 2026 scaling", "Momentum (£59k)"]
                ],
                { columnWidths: [Math.floor(CONTENT_WIDTH * 0.65), Math.floor(CONTENT_WIDTH * 0.35)], highlightRow: 1 }
            ),

            // ========== PAGE 13: SCARCITY ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Availability", 1),

            createCalloutBox(
                "LIMITED AVAILABILITY",
                [
                    new Paragraph({
                        spacing: { after: 160 },
                        children: [new TextRun({ text: "We take on a maximum of 2 Foundation engagements per quarter to ensure every client gets our full attention and expertise.", font: "Calibri", size: 22, color: TEXT_DARK })]
                    }),
                    new Paragraph({
                        spacing: { after: 100 },
                        children: [new TextRun({ text: "Current Q1 2026 Status:", bold: true, font: "Calibri", size: 22, color: TEXT_DARK })]
                    }),
                    createBullet("Slot 1: Booked (engagement starts late February)"),
                    new Paragraph({
                        numbering: { reference: "bullets", level: 0 },
                        spacing: { after: 100 },
                        children: [
                            new TextRun({ text: "Slot 2: ", font: "Calibri", size: 22, color: TEXT_DARK }),
                            new TextRun({ text: "AVAILABLE", bold: true, font: "Calibri", size: 22, color: SUCCESS_GREEN }),
                            new TextRun({ text: ". This is the slot we're offering you", font: "Calibri", size: 22, color: TEXT_DARK })
                        ]
                    }),
                    new Paragraph({
                        spacing: { before: 160 },
                        children: [new TextRun({ text: "If this slot is not confirmed by February 28, it will be offered to another client on our waitlist.", italics: true, font: "Calibri", size: 22, color: DARK_GRAY })]
                    })
                ],
                LIGHT_GRAY,
                NAVY_DARK,
                NAVY_DARK
            ),

            spacer(300),

            createHeading("Bonus Expiry", 2),
            createMultiParagraph([
                { text: "The 90-Day Health Check and Team Onboarding Template Pack bonuses (" },
                { text: "£4,000 combined value", bold: true },
                { text: ") are only available if you confirm your chosen option by " },
                { text: "February 28, 2026.", bold: true }
            ]),
            createParagraph("After this date, these bonuses are removed and the core engagement price remains the same."),

            // ========== PAGE 14: LOGISTICS ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Logistics & Access", 1),

            createParagraph("All engagement options require some level of access for team interviews and (for Foundation/Momentum) implementation work."),

            createHeading("Access Requirements", 2),
            createBullet("Team interviews: Platform team (3 FTEs), 2 onboarded application teams, 1-2 teams awaiting onboarding"),
            createBullet("For Foundation/Momentum: Access to platform infrastructure for implementation"),
            createBullet("On-site presence in Switzerland for interview and implementation phases"),

            spacer(200),

            createHeading("Classified Environment", 2),
            createParagraph("We understand Weisshorn operates in a classified environment. Platform Fix is prepared to:"),
            createBullet("Sign your NDA (Swiss-compliant) prior to engagement"),
            createBullet("Work within your security protocols and access controls"),
            createBullet("Ensure all findings and documentation remain confidential to Weisshorn"),
            createBullet("No platform data leaves your environment"),

            spacer(100),

            createParagraph("If direct access is not feasible, we can discuss an alternative approach involving training internal champions to gather assessment data, with remote review and advisory. This is not the preferred approach as it limits our ability to deliver on our guarantees, but we can explore if needed."),

            spacer(200),

            createHeading("Travel Arrangements", 2),

            createEliteTable(
                ["Responsibility", "Arrangement"],
                [
                    ["Client pays directly", "Hotel accommodation"],
                    ["Platform Fix invoices at cost", "Return flights (UK to Switzerland)"],
                    ["Platform Fix invoices at cost", "Meals and incidentals (with receipts)"]
                ],
                { columnWidths: [3500, CONTENT_WIDTH - 3500] }
            ),

            spacer(200),

            createEliteTable(
                ["Engagement", "Estimated Travel Cost"],
                [
                    ["Audit (1-2 days on-site)", "~£800"],
                    ["Foundation/Momentum (2 weeks on-site)", "~£2,500"]
                ],
                { columnWidths: [Math.floor(CONTENT_WIDTH * 0.6), Math.floor(CONTENT_WIDTH * 0.4)] }
            ),

            // ========== PAGE 15: INVESTMENT ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Investment Summary", 1),

            createEliteTable(
                ["Option", "Investment", "Total Value", "You Save", "Travel"],
                [
                    ["Platform Audit", "£7,500", "£7,500", "—", "~£800"],
                    [{ text: "Foundation ★", bold: true }, { text: "£35,000", bold: true }, { text: "£48,500", bold: true }, { text: "£13,500 (28%)", bold: true, color: SUCCESS_GREEN }, "~£2,500"],
                    ["Momentum", "£59,000", "£67,000", "£8,000 (12%)", "~£2,500"]
                ],
                { columnWidths: [2200, 1800, 1800, 2000, 1706], highlightRow: 1 }
            ),

            spacer(400),

            createHeading("Payment Terms", 1),

            createEliteTable(
                ["Option", "On Signing", "On Completion"],
                [
                    ["Audit (£7.5k)", "£7,500 (100%)", "—"],
                    [{ text: "Foundation (£35k)", bold: true }, { text: "£12,500", bold: true }, { text: "£22,500 (Week 6)", bold: true }],
                    ["Momentum (£59k)", "£12,500", "£22,500 (Week 6) + £4k/mo × 6"]
                ],
                { columnWidths: [3000, 3000, 3506], highlightRow: 1 }
            ),

            spacer(200),

            createParagraph("Travel costs invoiced separately at cost with receipts."),

            // ========== PAGE 16: NEXT STEPS ==========
            new Paragraph({ children: [new PageBreak()] }),

            createHeading("Next Steps", 1),

            createNumberedItem("You confirm classified environment access for team interviews"),
            createNumberedItem("You select an option (or we discuss further)"),
            createNumberedItem("We arrange NDA signing (your standard Swiss-compliant NDA is fine)"),
            createNumberedItem("I send the formal Statement of Work for your chosen option"),
            createNumberedItem("We book kickoff (targeting late February / early March 2026)"),

            spacer(200),

            createParagraph("I'm available to discuss any questions or concerns. Let's schedule a follow-up call for w/c 9th February to review and agree next steps."),

            spacer(400),

            // Validity box
            createCalloutBox(
                "PROPOSAL VALIDITY",
                [
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 100 },
                        children: [new TextRun({ text: "This proposal is valid until February 28, 2026.", bold: true, font: "Calibri", size: 24, color: TEXT_DARK })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: "After this date: Pricing and availability subject to change • Bonuses (£4,000 value) removed • Q1 slot may be allocated elsewhere", font: "Calibri", size: 20, color: DARK_GRAY })]
                    })
                ],
                GOLD_LIGHT,
                GOLD,
                GOLD
            ),

            spacer(400),

            // Contact box
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                columnWidths: [CONTENT_WIDTH],
                rows: [new TableRow({
                    children: [new TableCell({
                        borders: navyBorders,
                        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
                        shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
                        margins: { top: 200, bottom: 200, left: 300, right: 300 },
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 60 },
                                children: [new TextRun({ text: "Questions?", font: "Calibri", size: 20, color: DARK_GRAY })]
                            }),
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 40 },
                                children: [new TextRun({ text: "Steve Wade", bold: true, font: "Calibri", size: 26, color: NAVY_DARK })]
                            }),
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 40 },
                                children: [new TextRun({ text: "Founder, Platform Fix", font: "Calibri", size: 20, color: DARK_GRAY })]
                            }),
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 40 },
                                children: [new TextRun({ text: "steve@platformfix.com", font: "Calibri", size: 20, color: NAVY_MID })]
                            }),
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                children: [new TextRun({ text: "platformfix.com", font: "Calibri", size: 20, color: NAVY_MID })]
                            })
                        ]
                    })]
                })]
            })
        ]
    }]
});

// Generate the document
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("/mnt/user-data/outputs/Weisshorn_C5I_Proposal_Elite.docx", buffer);
    console.log("Elite proposal document created successfully!");
});