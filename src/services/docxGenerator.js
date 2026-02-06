import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, WidthType, AlignmentType, BorderStyle, HeadingLevel,
  PageBreak,
} from 'docx';

const FONT = 'Calibri';
const BORDER = { style: BorderStyle.SINGLE, size: 1, color: '000000' };
const CELL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

function headerCell(text, widthPct) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    borders: CELL_BORDERS,
    shading: { fill: '4472C4' },
    children: [new Paragraph({
      children: [new TextRun({ text, bold: true, color: 'FFFFFF', font: FONT, size: 20 })],
      spacing: { before: 40, after: 40 },
    })],
  });
}

function cell(text, widthPct) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    borders: CELL_BORDERS,
    children: [new Paragraph({
      children: [new TextRun({ text: text || '', font: FONT, size: 20 })],
      spacing: { before: 40, after: 40 },
    })],
  });
}

function sectionHeading(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text: `${number} ${title}`, bold: true, font: FONT, size: 24, color: '2F5496' })],
    spacing: { before: 300, after: 120 },
  });
}

export async function generatePMDocument(data) {
  const today = new Date().toLocaleDateString('en-IN');

  // --- Approval table ---
  const approvalTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [headerCell('', 20), headerCell('Prepared By / Last Updated By', 27), headerCell('Reviewed By', 27), headerCell('Approved By', 26)] }),
      new TableRow({ children: [cell('Name', 20), cell(data.preparedByName || '', 27), cell(data.reviewedByName || '', 27), cell(data.approvedByName || '', 26)] }),
      new TableRow({ children: [cell('Role', 20), cell(data.preparedByRole || '', 27), cell(data.reviewedByRole || '', 27), cell(data.approvedByRole || '', 26)] }),
      new TableRow({ children: [cell('Signature', 20), cell('', 27), cell('', 27), cell('', 26)] }),
      new TableRow({ children: [cell('Date', 20), cell(data.preparedByDate || today, 27), cell(data.reviewedByDate || today, 27), cell(data.approvedByDate || today, 26)] }),
    ],
  });

  // --- Expected System Impact table ---
  const impactRows = (data.systemImpacts || []).map(row =>
    new TableRow({ children: [cell(row.application, 25), cell(row.components, 35), cell(row.remarks, 40)] })
  );
  const systemImpactTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [headerCell('Impacted Application(s)', 25), headerCell('Impacted File(s) / Component(s) Name', 35), headerCell('Remarks', 40)] }),
      ...(impactRows.length > 0 ? impactRows : [new TableRow({ children: [cell('', 25), cell('', 35), cell('', 40)] })]),
    ],
  });

  // --- Assumptions and Risk table ---
  const riskRows = (data.risks || []).map(row =>
    new TableRow({ children: [cell(row.assumptions, 25), cell(row.risks, 25), cell(row.otherImpacts, 25), cell(row.remarks, 25)] })
  );
  const riskTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [headerCell('Assumptions', 25), headerCell('Risk(s)', 25), headerCell('Other Impact(s)', 25), headerCell('Remarks', 25)] }),
      ...(riskRows.length > 0 ? riskRows : [new TableRow({ children: [cell('', 25), cell('', 25), cell('', 25), cell('', 25)] })]),
    ],
  });

  // --- Impact Analysis Details table ---
  const detailRows = (data.analysisDetails || []).map(row =>
    new TableRow({ children: [
      cell(row.application, 18), cell(row.components, 22),
      cell(row.regressionNeeded || 'No', 12), cell(row.dbImpact || 'No', 12),
      cell(row.effort || '', 12), cell(row.remarks, 24),
    ]})
  );
  const detailsTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [
        headerCell('Impacted Application(s)', 18), headerCell('Impacted File(s) / Component(s) Name', 22),
        headerCell('Regression Testing Needed (Yes/No)', 12), headerCell('Database Impact (Yes/No)', 12),
        headerCell('Estimated Effort (Person-hours)', 12), headerCell('Remarks', 24),
      ]}),
      ...(detailRows.length > 0 ? detailRows : [new TableRow({ children: [cell('', 18), cell('', 22), cell('', 12), cell('', 12), cell('', 12), cell('', 24)] })]),
    ],
  });

  // --- Change Log table ---
  const changeLogTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: [headerCell('Version Number', 15), headerCell('Changes Made', 30), headerCell('Section No.', 10), headerCell('Changed By', 15), headerCell('Effective Date', 15), headerCell('Changes Effected', 15)] }),
      new TableRow({ children: [
        cell(data.versionNumber || 'V1.0.0', 15),
        cell(data.changesMade || 'Initial baseline created', 30),
        cell('', 10),
        cell(data.changedBy || data.preparedByName || '', 15),
        cell(data.effectiveDate || today, 15),
        cell('', 15),
      ]}),
    ],
  });

  const pmNumber = data.pmNumber || '';
  const crNumber = data.crNumber || '';

  const doc = new Document({
    sections: [{
      properties: {},
      children: [
        // Title
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '<SBIG>', bold: true, font: FONT, size: 28, color: '2F5496' })],
          spacing: { after: 100 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: `${pmNumber ? `PM${pmNumber} ` : 'PM '}Impact Analysis Document`, bold: true, font: FONT, size: 32, color: '2F5496' })],
          spacing: { after: 40 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: `<Version No.${data.versionNumber || '1'}>`, font: FONT, size: 20, color: '666666' })],
          spacing: { after: 200 },
        }),

        // Approval table
        approvalTable,
        new Paragraph({ children: [new TextRun({ text: '' })], spacing: { after: 100 } }),

        // Table of Contents
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun({ text: 'Table of Contents', bold: true, font: FONT, size: 24, color: '2F5496' })],
          spacing: { before: 200, after: 100 },
        }),
        new Paragraph({ children: [new TextRun({ text: '1.0  Introduction', font: FONT, size: 20 })] }),
        new Paragraph({ children: [new TextRun({ text: '2.0  Expected System Impact', font: FONT, size: 20 })] }),
        new Paragraph({ children: [new TextRun({ text: '3.0  Assumptions and Risk / Other Impact', font: FONT, size: 20 })] }),
        new Paragraph({ children: [new TextRun({ text: '4.0  Impact Analysis Details', font: FONT, size: 20 })] }),
        new Paragraph({ children: [new TextRun({ text: '5.0  Change Log', font: FONT, size: 20 })], spacing: { after: 200 } }),

        // 1.0 Introduction
        sectionHeading('1.0', 'Introduction'),
        new Paragraph({
          children: [new TextRun({ text: data.issueDescription || '', font: FONT, size: 20 })],
          spacing: { after: 80 },
        }),
        new Paragraph({
          children: [
            new TextRun({ text: 'Issue Description: ', bold: true, font: FONT, size: 20 }),
            new TextRun({ text: data.issueDescriptionDetail || data.issueDescription || '', font: FONT, size: 20 }),
          ],
          spacing: { after: 200 },
        }),

        // 2.0 Expected System Impact
        sectionHeading('2.0', 'Expected System Impact'),
        systemImpactTable,
        new Paragraph({ children: [new TextRun({ text: '' })], spacing: { after: 200 } }),

        // 3.0 Assumptions and Risk
        sectionHeading('3.0', 'Assumptions and Risk / Other Impact'),
        riskTable,
        new Paragraph({ children: [new TextRun({ text: '' })], spacing: { after: 200 } }),

        // 4.0 Impact Analysis Details
        sectionHeading('4.0', 'Impact Analysis Details'),
        detailsTable,
        new Paragraph({ children: [new TextRun({ text: '' })], spacing: { after: 200 } }),

        // 5.0 Change Log
        sectionHeading('5.0', 'Change Log'),
        changeLogTable,
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  const fileName = pmNumber ? `PM${pmNumber} Impact Analysis Document.docx` : 'PM Impact Analysis Document.docx';
  return { blob, fileName };
}
