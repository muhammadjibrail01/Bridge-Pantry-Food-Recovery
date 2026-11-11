// Builds outputs/Bridge_Pantry_Research_Paper.docx from the model results and figures.
// Run after src/model.py and src/charts.py. The contents page uses page numbers in
// outputs/toc_pages.json, which src/toc.py refreshes from a rendered PDF.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  HeadingLevel, BorderStyle, ShadingType, Footer, Header, PageNumber, PageBreak, TabStopType, Tab
} = require('docx');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(ROOT, 'outputs');
const FIG = path.join(OUT, 'figures');
const MODEL = JSON.parse(fs.readFileSync(path.join(OUT, 'model_out.json')));
const T = MODEL.truck_base, S1 = MODEL.S1, SP = MODEL.spons;

// Food Depository benchmarks (data/processed/gcfd_benchmarks.csv; the last column may contain commas)
const GCFD = {};
fs.readFileSync(path.join(ROOT, 'data', 'processed', 'gcfd_benchmarks.csv'), 'utf8').trim().split('\n').slice(1).forEach(line => {
  const parts = line.split(',');
  GCFD[parts[0]] = [1, 2, 3].map(i => (parts[i] === '' ? null : Number(parts[i])));
});

// Number formatting used in tables
const fmtInt = v => Math.round(v).toLocaleString('en-US');
const fmtUSD = v => (v < 0 ? '\u2212$' : '$') + Math.abs(Math.round(v)).toLocaleString('en-US');
const fmtUSD2 = v => '$' + v.toFixed(2);
const fmtM = v => (v < 0 ? '\u2212$' : '$') + (Math.abs(v) / 1e6).toFixed(2) + 'M';
const fmtMeals = v => (v / 1e6).toFixed(1) + 'M';
const pct0 = v => Math.round(v * 100) + '%';
const round100 = v => Math.round(v / 100) * 100;

const FONT = 'Times New Roman';
const BODY_SIZE = 24;      // 12 pt
const SMALL = 18;          // 9 pt
const TBL = 19;            // 9.5 pt

// ---------- inline markup: *italic*  **bold** ----------
function runs(text, opts = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size: opts.size || BODY_SIZE, bold: opts.bold, italics: opts.italics }));
    const tok = m[0];
    if (tok.startsWith('**')) out.push(new TextRun({ text: tok.slice(2, -2), font: FONT, size: opts.size || BODY_SIZE, bold: true, italics: opts.italics }));
    else out.push(new TextRun({ text: tok.slice(1, -1), font: FONT, size: opts.size || BODY_SIZE, italics: true, bold: opts.bold }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, size: opts.size || BODY_SIZE, bold: opts.bold, italics: opts.italics }));
  return out;
}
const P = (text, o = {}) => new Paragraph({
  children: runs(text, o), alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after ?? 140, line: o.line || 276 }, indent: o.noIndent ? undefined : { firstLine: 360 }
});
const PN = (text, o = {}) => P(text, { ...o, noIndent: true });
const HEADS = [];
const H1 = (t) => (HEADS.push([1, t]), new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })], spacing: { before: 300, after: 140 } }));
const H2 = (t) => (HEADS.push([2, t]), new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })], spacing: { before: 220, after: 100 } }));
const Note = (text) => new Paragraph({ children: runs(text, { size: SMALL }), spacing: { after: 200, line: 240 }, alignment: AlignmentType.LEFT });
const Cap = (text) => new Paragraph({ children: runs(text, { size: 21, bold: true }), spacing: { before: 160, after: 80, line: 240 }, keepNext: true, alignment: AlignmentType.LEFT });
const PB = () => new Paragraph({ children: [new PageBreak()] });

function Fig(file, wIn, hIn, caption, note) {
  const img = new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 60 }, keepNext: true,
    children: [new ImageRun({ type: 'png', data: fs.readFileSync(path.join(FIG, file)), transformation: { width: Math.round(wIn * 96), height: Math.round(hIn * 96) },
      altText: { title: caption, description: caption, name: file } })]
  });
  return [Cap(caption), img, Note(note)];
}

const border = { style: BorderStyle.SINGLE, size: 4, color: '808080' };
const noB = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
function Tbl(headers, rows, widths, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const mk = (t, isHead, i, ri) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA },
    borders: { top: isHead ? border : noB, bottom: isHead || ri === rows.length - 1 ? border : { style: BorderStyle.SINGLE, size: 2, color: 'D0D0D0' }, left: noB, right: noB },
    shading: isHead ? { fill: 'E7ECF2', type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: [new Paragraph({ alignment: (i === 0 || opts.leftAll) ? AlignmentType.LEFT : AlignmentType.RIGHT, spacing: { line: 240 },
      children: runs(String(t), { size: TBL, bold: isHead || (opts.boldLast && ri === rows.length - 1) }) })]
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => mk(h, true, i, -1)) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => mk(c, false, i, ri)) }))]
  });
}
const Ref = (text) => new Paragraph({ children: runs(text, { size: 22 }), spacing: { after: 100, line: 240 }, indent: { left: 540, hanging: 540 }, alignment: AlignmentType.LEFT });

// ======================= CONTENT =======================
const C = [];

// ---------- Title page ----------
C.push(new Paragraph({ spacing: { before: 2200, after: 200 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'The Bridge Pantry', font: FONT, size: 44, bold: true })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 },
  children: [new TextRun({ text: 'A Financially Sustainable Model for Reducing Food Waste and Homelessness in Chicago', font: FONT, size: 30 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 900 },
  children: [new TextRun({ text: 'From Food Waste to Social Rehabilitation: A Public-Private Partnership Model for Addressing Hunger and Homelessness in Chicago', font: FONT, size: 24, italics: true })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: 'Muhammad Jibrail', font: FONT, size: 26 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: 'Harris School of Public Policy, The University of Chicago', font: FONT, size: 24 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1400 }, children: [new TextRun({ text: 'October 2026', font: FONT, size: 24 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: 'The Bridge Pantry Chicago is a fiscally sponsored project of the Greater Chatham Initiative, a 501(c)(3) nonprofit organization.', font: FONT, size: 20, italics: true })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Citation style: The Chicago Manual of Style, author-date system.', font: FONT, size: 20, italics: true })] }));
C.push(PB());

// ---------- Abstract ----------
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [new TextRun({ text: 'Abstract', font: FONT, size: 26, bold: true })] }));
C.push(PN('Chicago discards prepared restaurant food every night while thousands of residents sleep unsheltered and more than a third of households report food insecurity. This paper asks whether a restaurant-to-street food recovery operation can be scaled into a financially sustainable social enterprise that reduces both food waste and the harms of homelessness. It uses the case of The Bridge Pantry, a volunteer initiative that has redistributed surplus quick-service restaurant meals to unhoused residents since January 2026 and now collects from eight Nando\u2019s locations. Drawing on ReFED\u2019s 2024 Illinois surplus estimates, USDA SNAP administrative data, Chicago\u2019s 2026 Point-in-Time count, audited financial statements of the Greater Chicago Food Depository, EPA life-cycle factors, and a review of Illinois and Chicago law, the paper builds a ten-year cost-benefit model of a sponsored refrigerated-truck program. Illinois foodservice businesses generated about 455,000 tons of surplus food in 2024 and donated roughly 0.24 percent of it. A base-case truck delivers 93,600 meals a year at $1.44 per meal in cash costs. A three-truck program yields a societal net present value of about $0.45 million at a 3 percent real discount rate (benefit-cost ratio 1.11), with a 78 percent probability of a positive return in Monte Carlo simulation. Sponsorship alone covers about 36 percent of operating costs, so the model is viable only as a blended-finance enterprise. The paper closes with a tiered Illinois edible-food recovery proposal, a SNAP-complementary framing, and a design for reintegration centers and social-enterprise caf\u00e9s.', { line: 264 }));
C.push(PN('*Keywords:* food recovery; homelessness; social enterprise; cost-benefit analysis; corporate social responsibility; SNAP; Illinois food policy', { line: 264 }));
C.push(PB());

// ---------- Contents (filled in below) ----------
C.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 160 }, children: [new TextRun({ text: 'Contents', font: FONT, size: 26, bold: true })] }));
const TOC_INDEX = C.length; C.push(null);
C.push(PB());

// ---------- Executive Summary ----------
C.push(H1('Executive Summary'));
C.push(P('The Bridge Pantry began in January 2026 as two volunteers, a borrowed car, and three to four boxes of unsold food collected each day from two quick-service restaurants and handed out under bridges downtown and in Hyde Park. Since then it has gained a fiscal sponsor, the Greater Chatham Initiative, and a formal partnership with eight Nando\u2019s PERi-PERi locations that now release surplus on a weekly evening route from Hyde Park to Lakeview. This paper evaluates whether that model can become a durable, financially sustainable institution, and what public policy would help it and organizations like it scale.'));
C.push(P('The need is not in doubt. Chicago counted 6,527 people in shelters or on the street on one night in January 2026, including 1,641 people who were unsheltered, about a quarter more than the year before (DFSS 2026). A broader measure that includes people doubled up with others puts the 2024 total at 58,625 (Nelsen, Gray, and Paler-Ponce 2026). About 36 percent of Chicago households were food insecure in June 2026 (GCFD 2026a), and Illinois SNAP enrollment fell 18.8 percent in the twelve months to June 2026 after federal changes to the program (USDA FNA 2026a). At the same time, Illinois restaurants, caterers, hotels, and other foodservice businesses generated an estimated 454,817 tons of surplus food in 2024, of which only 1,072 tons, or 0.24 percent, were donated; nearly all of the rest was landfilled (ReFED 2026a).'));
C.push(P('Four findings stand out. First, a single sponsored refrigerated truck running six nights a week can plausibly deliver about 93,600 meals a year at a cash cost of $1.44 per meal ($1.63 including capital), which is below the $2.07 per meal that SNAP currently provides an Illinois participant and well below the $3.70 national average cost of a meal (Feeding America 2026a). Second, the advertising-only financing model in the original proposal does not close: a $4,000 monthly sponsorship covers about 36 percent of a truck\u2019s operating cost, and full self-sufficiency would require roughly $9,100 per month, about 35 percent of the market value of a dedicated LED billboard truck on comparable hours. The program is therefore best designed as a blended-finance social enterprise combining sponsorship, restaurant CSR partnerships, grants, and individual giving. Third, the program is socially worthwhile under conservative assumptions: the ten-year societal net present value is about $0.45 million at a 3 percent real discount rate, the result remains positive at 7 percent, and the program breaks even as long as each consumed meal is worth at least $1.79 to its recipient. Fourth, the climate case is weaker than commonly claimed; under conservative accounting, landfill methane avoided by one truck is almost exactly offset by its diesel emissions, so the environmental benefit depends on whether rescued meals displace new food production.'));
C.push(P('The paper recommends a phased path: a lean Stage 2 pilot using a used refrigerated van and volunteer drivers; a formal partnership with the Greater Chicago Food Depository rather than parallel infrastructure; a tiered Illinois edible-food recovery requirement for large generators modeled on California and New York; alignment of Illinois donor-liability law with the 2022 federal Food Donation Improvement Act; expansion of the SNAP Restaurant Meals Program in Chicago as a complement to food recovery; and, once operations are stable, reintegration centers and a training caf\u00e9 built on evidence from Chicago\u2019s Inspiration Kitchens and the REDF social enterprise evaluations.'));

// ---------- 1 Introduction ----------
C.push(H1('1. Introduction'));
C.push(P('Food waste and homelessness are usually treated as separate problems with separate constituencies. Food waste belongs to environmental agencies, sustainability officers, and waste haulers; homelessness belongs to shelters, outreach teams, and housing authorities. Yet on any evening in downtown Chicago the two problems sit a few blocks apart. Restaurants close with trays of safe, prepared food that will be thrown away within the hour, while people sleeping under Lower Wacker Drive or along the river go without a hot meal. The Bridge Pantry was founded to close that gap with the most direct intervention available: collect the food before it is discarded and deliver it the same night.'));
C.push(P('Direct food rescue is an old idea, but scaling it is hard. Volunteer operations run on goodwill, private vehicles, and the founders\u2019 own money, and they rarely survive the departure of their founders. Large food banks have the logistics but are organized around pantries, shelf-stable inventory, and fixed distribution sites rather than hot meals delivered to people living outside. This paper therefore asks a practical policy question: *Can a restaurant-to-street food recovery system be scaled into a financially self-sustaining enterprise that simultaneously reduces food waste and the harms of homelessness, and what public policies would make that possible?*'));
C.push(P('The paper makes three contributions. It documents the operating model of a new Chicago food recovery initiative and quantifies its current scale. It builds a transparent ten-year cost-benefit and financial model of the proposed sponsored-truck expansion, including a valuation of the advertising concept at the heart of the original proposal and a Monte Carlo analysis of uncertainty. And it connects that operational model to Illinois policy, proposing specific legal and programmatic changes grounded in comparisons with other states and in current federal law. The analysis is deliberately critical of the original proposal where the evidence requires it, because a plan that survives scrutiny is more useful to funders, partners, and legislators than one that does not.'));
C.push(P('Section 2 describes the Chicago context and The Bridge Pantry. Section 3 reviews the relevant literature. Section 4 sets out data and methods. Sections 5 through 7 analyze Stage 1 operations, the Stage 2 sponsored-truck model, and the cost-benefit results. Sections 8 and 9 present the Illinois policy proposal and the SNAP framing. Section 10 designs Stage 4 reintegration services, and Sections 11 through 13 cover measurement, implementation, and conclusions.'));
C.push(...Fig('fig1_stages.png', 6.4, 2.2, 'Figure 1. The four-stage Bridge Pantry model', 'Source: Author.'));

// ---------- 2 Background ----------
C.push(H1('2. Background'));
C.push(H2('2.1 Homelessness in Chicago'));
C.push(P('Chicago\u2019s official measure of homelessness is the annual Point-in-Time (PIT) count, a one-night census of people in shelters, transitional housing, and places not meant for habitation, conducted under U.S. Department of Housing and Urban Development (HUD) rules. The 2026 count, conducted on January 22 during a wind chill of minus 35 degrees, found 6,527 people: 4,886 in sheltered locations and 1,641 unsheltered (DFSS 2026). The total is down sharply from the 2024 peak of 18,836, but that peak was driven almost entirely by asylum-seeking new arrivals housed in city shelters. Excluding new arrivals and newcomers, the count fell from 6,786 in 2015 to a low of 3,875 in 2022 and has since climbed back to 6,527, its highest level since 2015, while unsheltered homelessness rose by about a quarter between 2025 and 2026 (Figure 2). The city notes that the extreme cold pushed many people temporarily indoors onto trains and into police stations, which may have increased the number counted (DFSS 2026).'));
C.push(P('Two features of the 2026 count matter directly for The Bridge Pantry. The unsheltered population is entirely composed of adult-only households, and 58 percent of unsheltered people were between 35 and 54 years old. And 68 percent of unsheltered people identified as Black, compared with less than a third of the city\u2019s population (DFSS 2026). A street-based meal program in the Loop and on the South Side is therefore serving a population that is older, adult, and disproportionately Black, and that is poorly reached by pantries designed for families with kitchens.'));
C.push(P('The PIT count is a floor, not a ceiling. Because HUD excludes people staying temporarily with friends or relatives, the Chicago Coalition to End Homelessness combines city administrative data with census microdata to estimate how many people experience homelessness over a full year. Its most recent estimate is 58,625 people in 2024, of whom 19,317 were in shelters or on the street and 39,308 were doubled up (Nelsen, Gray, and Paler-Ponce 2026; see also Richard et al. 2022 for the method). The two measures answer different questions, and both are used below: the PIT count to size the nightly unsheltered population that a mobile meal program reaches directly, and the annual estimate to size the population a reintegration center might eventually serve.'));
C.push(...Fig('fig2_pit.png', 6.3, 2.9, 'Figure 2. Chicago Point-in-Time count, 2015\u20132026', 'Source: DFSS (2026). "New arrivals" are asylum seekers who arrived between August 2022 and January 2024; "newcomers" arrived between January 2024 and January 2025.'));

C.push(H2('2.2 Food Insecurity and the Retreat of SNAP'));
C.push(P('Food insecurity in Chicago is high and rising. A representative survey of 1,194 Chicago residents conducted by NORC at the University of Chicago for the Greater Chicago Food Depository found that 36 percent of households, about 417,000, were food insecure in June 2026, and 19 percent had very low food security, the most severe category (GCFD 2026a). Rates reached 51 percent on the South Side, 47 percent among households with children, and 61 percent among Black households. Most adults in food-insecure households were working (56 percent). The same report attributes the pressure to a 22 percent rise in food prices and a 31 percent rise in rents in the Chicago area since June 2021.'));
C.push(P('The Supplemental Nutrition Assistance Program (SNAP) remains the main public defense against hunger, and Illinois has historically used it well: in fiscal year 2022 the state\u2019s participation rate among eligible people was estimated at or near 100 percent, compared with 88 percent nationally (Cunnyngham 2025). That changed after the July 2025 federal budget reconciliation law (H.R. 1), which the Food Depository estimates cuts SNAP by $186 billion over ten years and which Feeding America describes as shifting program costs to the states (GCFD 2026a; Feeding America 2026a). Illinois participation fell from 1.87 million people in June 2025 to 1.52 million in June 2026, a drop of 350,633 people, and monthly benefits fell 21 percent to $287 million (USDA FNA 2026a). Because Illinois already enrolled nearly everyone who was eligible, the decline most likely reflects tighter eligibility rather than people failing to apply. The average benefit, about $189 per person per month, equals roughly $2.07 per meal.'));
C.push(...Fig('fig3_snap.png', 6.3, 2.6, 'Figure 3. Illinois SNAP participation, FY2005\u2013FY2026', 'Source: Author\u2019s calculations from USDA FNA (2026a; 2026b). FY2026 covers October 2025 through June 2026.'));

C.push(H2('2.3 Surplus Food in Illinois'));
C.push(P('ReFED\u2019s Insights Engine, the most detailed public model of U.S. food surplus, estimates that Illinois generated 2.16 million tons of surplus food in 2024 across farms, manufacturers, retailers, foodservice, and households, enough for an estimated 3.5 billion meals (ReFED 2026a). The foodservice sector, which includes restaurants, caterers, schools, hotels, and hospitals, accounted for 454,817 tons with a retail value of about $5.6 billion. Only 1,072 tons of foodservice surplus were donated, about 0.24 percent, and 99.7 percent was landfilled (Figure 4). Limited-service restaurants, the segment that includes Nando\u2019s and the project\u2019s other quick-service sources, generated 110,622 tons and donated 0.48 percent. Chicken-focused limited-service restaurants alone generated an estimated 11,131 tons.'));
C.push(P('Not all of that surplus is recoverable. Nationally, plate waste, food left on customers\u2019 plates, accounts for 36 percent of limited-service restaurant surplus and cannot be donated. Overproduction and catering overproduction, which can be, account for 25 and 19 percent respectively (ReFED 2026a). ReFED\u2019s methodology also notes that commercial kitchens leave about 4.2 percent of purchased food unused, and that roughly 48 percent of that pre-consumer surplus at limited-service restaurants is overproduction (ReFED 2026b). The recoverable share of restaurant surplus is therefore large in absolute terms but modest relative to total waste, and it is concentrated exactly where The Bridge Pantry operates: cooked, ready-to-eat food at the end of the evening.'));
C.push(...Fig('fig4_refed.png', 6.3, 2.8, 'Figure 4. Illinois foodservice surplus by segment and share donated, 2024', 'Source: Author\u2019s calculations from ReFED (2026a). Donated tonnage is too small to be visible at this scale.'));

C.push(H2('2.4 The Bridge Pantry'));
C.push(P('The Bridge Pantry was founded by students at the University of Chicago\u2019s Harris School of Public Policy. From January 2026 it collected three to four boxes of unsold food each day from two quick-service restaurants, a fried-chicken chain and a doughnut shop, neither of which is a formal partner, and distributed the food from a volunteer\u2019s personal car to unhoused residents under bridges downtown, in Hyde Park, and in neighborhoods adjacent to the Loop. The founders paid all costs themselves until April 2026, when the Greater Chatham Initiative, a Chicago 501(c)(3), became the project\u2019s fiscal sponsor. The founders have requested surplus from University of Chicago dining on a daily basis, but the university is not a formal partner; notably, the University is already listed among the Greater Chicago Food Depository\u2019s food rescue donors for fiscal year 2025 (GCFD 2025a).'));
C.push(P('The most significant development is a formal partnership with eight Nando\u2019s PERi-PERi restaurants. Under the chain\u2019s Harvest Program, The Bridge Pantry collects surplus on a fixed Friday evening route with half-hour pickup windows between 6:00 and 10:00 p.m.: Hyde Park, South Loop, Michigan Avenue, Wabash, Streeterville, West Loop, Clybourn, and Lakeview (The Bridge Pantry 2026). The partnership gives the project a predictable, branded supply chain across the central city and a natural first sponsor for the truck model in Stage 2. The executive proposal shared with Nando\u2019s asked each location for eight to ten meals or one tray per pickup, a named staff contact, sealed containers, and a 30- to 45-minute pickup window, at no cost to the restaurant (The Bridge Pantry 2025).'));

// ---------- 3 Literature ----------
C.push(H1('3. Literature Review'));
C.push(H2('3.1 Food Waste and the Environmental Case for Recovery'));
C.push(P('The U.S. Environmental Protection Agency replaced its long-standing Food Recovery Hierarchy in 2023 with the Wasted Food Scale, which ranks eleven pathways by environmental impact using a review of life-cycle studies (US EPA 2023a). Source reduction ranks first, followed by donation and upcycling, because these are the only pathways that can displace new food production. Landfilling ranks near the bottom because food decomposes quickly into methane. EPA cautions, however, that donation is not environmentally free: collection and refrigerated distribution consume energy, and between 4 and 40 percent of donated food may itself go uneaten, depending on the study (US EPA 2023a; Alexander and Smaje 2008). The agency\u2019s Waste Reduction Model (WARM) quantifies the difference: landfilling a short ton of mixed food waste emits about 0.50 metric tons of carbon dioxide equivalent, while preventing a ton of poultry from being produced avoids about 2.45 tons (US EPA 2023b). These factors frame the climate analysis in Section 7 and explain why the result depends heavily on what happens to food demand when rescued meals are eaten.'));
C.push(H2('3.2 Charitable Food Assistance and Its Critics'));
C.push(P('A substantial literature questions whether charitable food recovery addresses hunger or merely manages it. Poppendieck\u2019s (1998) study of the American emergency food system argued that its growth coincided with, and helped legitimize, the retrenchment of public entitlements. Tarasuk and Eakin\u2019s (2005) ethnography of Canadian food banks showed that handling unsaleable industry surplus is labor-intensive and depends on volunteers, on recipients\u2019 need, and on recipients\u2019 lack of rights within the system. A critical synthesis by McIntyre et al. (2016) concluded that food banks are an inadequate response to household food insecurity, and Galli, Cavicchi, and Brunori (2019) modeled the vulnerability of assistance systems that depend on surplus: as retailers reduce waste, the charitable supply shrinks. Lohnes (2021) traces how donor-liability and tax laws have shaped surplus into a regulated charitable commodity.'));
C.push(P('This critique shapes the design choices in this paper in three ways. The Bridge Pantry is positioned as a complement to SNAP and to housing, not a substitute for either. Its reliance on volunteer labor is treated as a cost and a risk, not a free input. And the policy proposal favors source reduction first and recovery second, consistent with EPA\u2019s scale.'));
C.push(H2('3.3 Food Rescue Models and Local Comparators'));
C.push(P('Operations research on food recovery emphasizes logistics: matching unpredictable supply with need, minimizing time-temperature risk, and routing vehicles efficiently (Lee et al. 2017; Bacon and Baker 2017). The dominant U.S. model is the regional food bank. The Greater Chicago Food Depository, a Feeding America member, distributed the equivalent of more than 96 million meals in fiscal year 2025 through more than 850 partner agencies, with total operating expenses of $244.2 million, of which $157.7 million was the value of donated food (GCFD 2025a). Its cash operating expenses of $86.5 million imply a cost of about $0.90 per meal, an important efficiency benchmark. Its own reports also show how the system is changing: donated food rose from 28 percent of supply in fiscal 2023 to 42 percent in fiscal 2025, and trucking and freight costs rose from $1.7 million to $2.8 million, while the share of food reaching shelters and mobile programs fell (Table 1).'));
C.push(P('The Food Depository\u2019s partner lists reveal a gap. Its food rescue partners are overwhelmingly grocers and retailers; McDonald\u2019s and Panera Bread appear for the first time in the most recent year, Starbucks is the one quick-service chain listed in all three years, and Chicago restaurant groups appear mainly as cause-related marketing partners rather than as food donors (GCFD 2023a; GCFD 2024a; GCFD 2025a). Prepared-food rescue from restaurants to unsheltered people is therefore a small, underserved niche even within a highly capable regional system. Federal rules reinforce the point: under the Illinois administration of The Emergency Food Assistance Program (TEFAP), only food banks may operate mobile distribution sites (IDHS 2024), which means a small organization cannot run federally supplied mobile routes on its own but can partner with a food bank that does.'));

C.push(Cap('Table 1. Greater Chicago Food Depository operating indicators, FY2023\u2013FY2025'));
C.push(Tbl(['Indicator', 'FY2023', 'FY2024', 'FY2025'], [
  ['Meals distributed (equivalent)', 'n/a', '101 million', '96+ million'],
  ['Total operating expenses', ...GCFD.total_operating_expenses.map(v => '$' + (v / 1e6).toFixed(1) + 'M')],
  ['Trucking, freight, and storage', ...GCFD.trucking_freight_storage.map(v => '$' + (v / 1e6).toFixed(2) + 'M')],
  ['Leased vehicles', ...GCFD.leased_vehicles.map(fmtInt)],
  ['Food supply: donated share', ...GCFD.donated_share_of_food.map(pct0)],
  ['Food to group homes and shelters', '3%', 'n/a', '2%'],
  ['Food to mobile programs', '5%', 'n/a', '2%'],
  ['Employees / volunteers', ...[0, 1, 2].map(i => fmtInt(GCFD.employees[i]) + ' / ' + fmtInt(GCFD.volunteers[i]))],
], [3600, 1900, 1900, 1900]));
C.push(Note('Sources: GCFD (2023a; 2024a; 2025a; 2025b) and Forms 990 for fiscal years 2023 and 2024 (GCFD 2023b; 2024b). "n/a" indicates the item was not reported in comparable form.'));

C.push(H2('3.4 Corporate Social Responsibility and Cause-Related Marketing'));
C.push(P('The Stage 2 revenue model rests on cause-related marketing, which Varadarajan and Menon (1988) defined as linking a firm\u2019s marketing to a contribution to a designated cause. Porter and Kramer (2011) argue more broadly that firms create durable value when social impact and business interest coincide. For restaurants, food donation fits this logic unusually well: it reduces disposal costs, can qualify for an enhanced federal tax deduction available to all business types (ReFED 2026c), produces visible community goodwill, and turns a waste stream into a story. Chicago\u2019s restaurant sector already practices cause-related marketing at scale; the Food Depository lists dozens of local restaurant groups as cause marketing partners (GCFD 2025a). The open question, addressed in Section 6, is how much a sponsor will pay for visibility on a truck that serves unhoused residents at night, and whether that payment can be structured without compromising the dignity of the people served or the organization\u2019s tax status.'));
C.push(H2('3.5 Social Enterprise, Employment, and Housing First'));
C.push(P('Stage 4 draws on two bodies of evidence. The first concerns employment social enterprises, businesses that hire people facing barriers to work. A Mathematica evaluation of REDF-funded enterprises in California found that social enterprise employment increased the likelihood of employment a year later by 19 percentage points relative to applicants who were not hired, and the share of workers in stable housing roughly tripled over the year (Rotz, Maxwell, and Dunn 2015; REDF 2015). Chicago\u2019s Inspiration Corporation operates Inspiration Kitchens, a full-service social enterprise restaurant that trains about 75 people a year; its restaurant revenue covers about 10 percent of the training program\u2019s operating costs, and it uses in-kind donations from local restaurants and grocers (Inspiration Corporation 2026). The lesson is that a caf\u00e9 is a training platform that needs subsidy, not a profit center.'));
C.push(P('The second concerns housing. Randomized evidence shows that Housing First programs, which provide permanent housing without requiring treatment or sobriety first, keep chronically homeless people housed at higher rates than treatment-first programs (Tsemberis, Gulcur, and Nakae 2004), and that supportive housing reduces public spending on shelters, hospitals, and jails enough to offset much of its cost (Culhane, Metraux, and Hadley 2002). Food insecurity itself is associated with worse health and higher health care use (Gundersen and Ziliak 2015). A food program cannot end homelessness, but it can be a reliable point of contact that connects people to housing systems that can.'));

// ---------- 4 Data and Methods ----------
C.push(H1('4. Data and Methods'));
C.push(P('The analysis combines four types of evidence. Program data come from the founders\u2019 operating reports, the Nando\u2019s Harvest Program route sheet, and the executive proposal shared with restaurant partners (The Bridge Pantry 2025; 2026). Contextual data come from Chicago\u2019s 2026 PIT count (DFSS 2026), the CCEH annual estimate (Nelsen, Gray, and Paler-Ponce 2026), the 2026 Hunger in Chicago survey (GCFD 2026a), Map the Meal Gap (Feeding America 2026a; 2026b), USDA SNAP administrative tables for fiscal years 2005 through 2026 (USDA FNA 2026a; 2026b), and ReFED\u2019s 2024 state-level surplus model (ReFED 2026a). Benchmark financial data come from the Food Depository\u2019s audited statements and Forms 990 for fiscal years 2023 through 2025. Legal analysis covers the Illinois Good Samaritan Food Donor Act, the Illinois Vehicle Code, the federal Bill Emerson Good Samaritan Food Donation Act as amended in 2022, Chicago\u2019s mobile food and sign rules, and the TEFAP distribution manual. Emission factors come from EPA\u2019s Waste Reduction Model, version 16 (US EPA 2023b; 2023c).'));
C.push(P('Stage 1 volumes are based on the founders\u2019 reports of boxes collected. They are converted to meals using 8 meal-equivalents per box and 10 meals per Nando\u2019s location per weekly pickup, consistent with the partner request, and to weight using 1.2 pounds per meal. The weight conversion matches the ratios implied by ReFED\u2019s Illinois limited-service restaurant estimates (1.21 pounds per meal) and the Food Depository\u2019s reporting (about 1.2 pounds per meal). Stage 2 capital and operating costs are planning estimates assembled from published vendor rate ranges, Chicago licensing fees, EIA fuel prices, the city minimum wage, and the Food Depository\u2019s fringe-benefit ratio. They are summarized in Appendix B and should be replaced with vendor quotes before procurement.'));
C.push(P('The cost-benefit analysis follows OMB Circular A-4 of 2003, which was reinstated in February 2025 when the 2023 revision was rescinded (OMB 2003; OMB 2025). Results are reported in constant 2026 dollars over a ten-year horizon at the circular\u2019s 3 and 7 percent real discount rates, with a sensitivity case at the 1.6 percent ten-year real Treasury rate from OMB Circular A-94, Appendix C (OMB 2026). The analysis takes the perspective of Chicago and Illinois society and measures the program incrementally against the Stage 1 status quo. Grants and donations are transfers and are excluded from benefits; sponsorship payments are treated as purchases of advertising services, so the sponsor\u2019s benefit is valued at the price paid, a lower bound on its willingness to pay. Volunteer time is costed at Chicago\u2019s July 2026 minimum wage of $17.05 per hour as an opportunity cost (City of Chicago 2026c). Health and public-service savings are not monetized because no credible local estimate exists for a meal program alone; instead the paper reports a break-even value. Uncertainty is analyzed with one-way sensitivity tests and a 10,000-draw Monte Carlo simulation using triangular distributions over the ranges in Appendix A.'));

// ---------- 5 Stage 1 ----------
C.push(H1('5. Stage 1: Current Operations and Scaling Food Recovery'));
C.push(P('Stage 1 is a collect-and-distribute model. Food is identified at the end of service, packed by restaurant staff in sealed containers, collected within a short window, sorted into hot, cooked-chilled, and uncooked items, and distributed the same evening, with additional items routed to community refrigerators and partner food banks (The Bridge Pantry 2025). Distribution sites are the places unhoused people actually are: underpasses and lower-level roadways downtown, transit corridors, and encampments in Hyde Park and areas adjacent to the Loop.'));
C.push(P('Table 2 summarizes current scale. The two informal restaurant sources supply about 336 meal-equivalents a week, or about 13,400 since January 2026, and the Nando\u2019s route adds about 80 per week. On an annual basis, the current operation moves about 21,600 meals and 13 short tons of food that would otherwise be landfilled, at a cash cost of roughly $8,200 a year, or $0.38 per meal, borne until recently by the founders. The low cash cost reflects about 2,200 hours of unpaid volunteer time, which is the binding constraint on growth.'));
C.push(Cap('Table 2. Stage 1 operating baseline (annualized)'));
C.push(Tbl(['Item', 'Value', 'Basis'], [
  ['Informal QSR sources', fmtInt(S1.informal_meals_week) + ' meals/week', '7 boxes/day \u00d7 6 days \u00d7 8 meals/box'],
  ['Nando\u2019s Harvest Program (8 locations)', fmtInt(S1.nandos_meals_week) + ' meals/week', '10 meals per location, Friday route'],
  ['Annual meals distributed', fmtInt(S1.meals_year), fmtInt(S1.meals_week) + ' meals/week \u00d7 52'],
  ['Food diverted from landfill', S1.tons_year.toFixed(1) + ' short tons', '1.2 lb per meal'],
  ['Vehicle miles', fmtInt(S1.miles), '35 miles/day, personal car'],
  ['Cash costs (fuel, packaging, misc.)', fmtUSD(S1.cash), 'Gasoline $4.15/gal; packaging $0.25/meal'],
  ['Cash cost per meal', fmtUSD2(S1.cash_per_meal), ''],
  ['Volunteer hours', fmtInt(S1.vol_hrs), '2 volunteers \u00d7 3.5 hours \u00d7 312 nights'],
], [3500, 2000, 3800], { leftAll: false }));
C.push(Note('Sources: Founders\u2019 operating reports; The Bridge Pantry (2026); EIA (2026b) for the Midwest gasoline price of October 12, 2026.'));
C.push(P('Three constraints limit Stage 1. The first is supply volatility: informal sources can stop at any time, and the Nando\u2019s route operates once a week, so nightly supply depends on the least formal partners. The second is cold-chain capacity: a passenger car cannot hold food at safe temperatures for long, which limits route length and the types of food accepted. The third is legal exposure. Illinois law protects restaurants and other food producers that donate prepared food in good faith to a nonprofit, and it protects the receiving nonprofit if it reasonably inspects the food at the time of donation and finds it apparently fit for consumption (745 ILCS 50/2.10, 2.11, 3, 4). The 2022 federal Food Donation Improvement Act adds a federal floor of protection and allows restaurants, as qualified direct donors, to donate directly to individuals (Food Donation Improvement Act of 2022). In practice, the protection depends on documentation, so a simple pickup log recording time, temperature, and condition is the single most valuable compliance step Stage 1 can take now. If the project passes food on to community refrigerators or other feeding programs, Illinois also requires food banks that supply such programs to file an annual list with the local health department for recall alerts (410 ILCS 625/3.2; IDPH n.d.).'));

// ---------- 6 Stage 2 ----------
C.push(H1('6. Stage 2: Financial Sustainability through Impact Marketing'));
C.push(H2('6.1 The Sponsored Truck'));
C.push(P('Stage 2 replaces the personal car with a refrigerated box truck that collects from partner restaurants on an evening route, holds food at safe temperatures, and distributes it at fixed nightly stops. The truck\u2019s exterior carries permanently mounted LED panels and a sponsor wrap, turning the vehicle into a moving acknowledgment of the restaurants and companies that fund it. The base case assumes six nights a week, 300 meals a night, a 55-mile route, a 16-foot diesel truck with a refrigeration unit, and a paid driver-coordinator working seven-hour shifts with one volunteer. The capital package is estimated at $140,000: $95,000 for the truck, $40,000 for the LED display system, and $5,000 for branding (Appendix B). Vendors list complete LED billboard trucks at about $50,000 to $150,000 to purchase (Can\u2019t Miss US 2025), and Chicago operators rent dedicated LED trucks at $1,200 to $1,400 per day (Boost Outdoor Media 2026).'));
C.push(Cap('Table 3. Annual operating budget per truck, base case (2026 dollars)'));
C.push(Tbl(['Line item', 'Annual cost', 'Basis'], [
  ['Diesel fuel (' + fmtInt(T.gal) + ' gallons)', fmtUSD(T.fuel), fmtInt(T.miles) + ' miles at 10 mpg + reefer; $5.00/gal'],
  ['Driver-coordinator wages and fringe', fmtUSD(T.labor), fmtInt(T.labor / (22 * 1.264)) + ' hours \u00d7 $22.00 \u00d7 1.264'],
  ['Commercial insurance', fmtUSD(T.insurance), 'City minimum $350,000 coverage'],
  ['Maintenance, reefer, and LED service', fmtUSD(T.maint), '$0.25/mile + $3,000 equipment'],
  ['Licenses, permits, GPS, food safety', fmtUSD(T.permits), 'Mobile food dispenser license $700/2 yrs'],
  ['Commissary and secure parking', fmtUSD(T.commissary), '$500 per month'],
  ['Food-safe packaging and supplies', fmtUSD(T.supplies), '$0.20 per meal'],
  ['Administration (12%)', fmtUSD(T.admin), 'Share of direct costs'],
  ['Fiscal sponsor fee (7% of contributions)', fmtUSD(T.fee), 'Pending sponsor agreement'],
  ['**Total cash operating cost**', '**' + fmtUSD(T.cash_cost) + '**', '**' + fmtUSD2(T.cost_per_meal_cash) + ' per meal**'],
], [3600, 1500, 4200]));
C.push(Note('Sources: Author\u2019s model; EIA (2026a); City of Chicago (2026a; 2026c); GCFD (2025a) for the fringe-benefit ratio. Capital cost of ' + fmtUSD(T.capex) + ' is annualized separately at ' + fmtUSD(T.annual_capital) + ' (seven-year life, 15 percent salvage).'));
C.push(H2('6.2 What Is a Sponsorship Worth?'));
C.push(P('The original proposal assumed $4,000 a month in advertising revenue per truck, or $480,000 a year for a ten-truck fleet. That figure can be tested against the market. A dedicated LED truck in Chicago costs $1,200 to $1,400 for an 8- to 10-hour day on a route the advertiser chooses (Boost Outdoor Media 2026). Pro-rated to a seven-hour evening shift, the Bridge Pantry truck offers about $1,000 of equivalent exposure per night, or roughly $26,000 a month over 26 operating nights. But a sponsor does not choose the route, much of the time is spent parked at distribution stops or traveling after dark, and the audience is partly residential. A discount of 60 to 90 percent from dedicated-truck prices is therefore realistic. At $4,000 a month, the base case prices the sponsorship at about 15 percent of dedicated-equivalent value, which appears achievable for a single anchor sponsor such as Nando\u2019s that values the association more than the impressions.'));
C.push(P('The problem is not the price but the arithmetic. As Table 4 shows, a $4,000 sponsorship covers about 36 percent of a truck\u2019s operating cost; with $20,000 in restaurant CSR contributions and $10,000 in individual giving, total contributed revenue covers 58 percent and leaves an annual gap of about $57,000 per truck. Full operating self-sufficiency would require about $9,100 a month in sponsorship, roughly 35 percent of dedicated-equivalent value, which in practice means selling two or three sponsor slots per truck. The ten-truck vision would cost about $1.35 million a year to operate against $480,000 in sponsorship. The model is viable, but as a blended-finance enterprise rather than an advertising business.'));
C.push(Cap('Table 4. Sponsorship scenarios per truck'));
C.push(Tbl(['Scenario', 'Sponsorship / month', 'Contributed revenue', 'Operating coverage', 'Annual gap', 'Program NPV (3%)'], [
  ...[['Low', SP.low], ['Base', SP.base], ['High (two to three slots)', SP.high]].map(([label, x]) =>
    [label, fmtUSD(x.sponsor_month), fmtUSD(x.contrib), pct0(x.coverage), fmtUSD(x.net), fmtM(x.npv3)]),
  ['Break-even', fmtUSD(MODEL.breakeven_sponsor), fmtUSD(12 * MODEL.breakeven_sponsor + T.csr + T.donations), '100%', '$0', '\u2014'],
], [2300, 1450, 1500, 1350, 1300, 1400]));
C.push(Note('Contributed revenue includes $20,000 in restaurant CSR contributions and $10,000 in individual donations per truck in all scenarios. Program NPV is the ten-year societal net present value of the three-truck program described in Section 7.'));
C.push(P('A leaner configuration narrows the gap considerably. Using a used refrigerated van, a smaller LED package, volunteer drivers, and a paid part-time coordinator, capital falls to about $85,000 and annual cash costs to about $91,000 for an estimated 68,640 meals, or $1.33 per meal. Contributed revenue then covers 86 percent of operating costs, and the program\u2019s ten-year societal NPV rises to about $0.65 million with a benefit-cost ratio of 1.22. The lean configuration is therefore the recommended Stage 2 pilot.'));
C.push(H2('6.3 Legal and Ethical Design Constraints'));
C.push(P('Four constraints shape how the truck and its sponsorships must be designed. First, **tax status.** Paid advertising sold by a 501(c)(3) can be taxable unrelated business income, while a qualified sponsorship payment, an acknowledgment of the sponsor\u2019s name and logo without comparative claims, prices, or calls to action, generally is not (26 U.S.C. \u00a7 513(i); 26 C.F.R. \u00a7 1.513-4). The Food Depository, for example, reports no advertising income on its unrelated business income return (GCFD 2025c). Sponsor content on the truck should therefore be acknowledgment-only, a requirement the fiscal sponsor\u2019s accountant should confirm before any agreement is signed.'));
C.push(P('Second, **vehicle law.** The Illinois Vehicle Code prohibits operating a vehicle with a video screen that is visible to the driver while driving, but exempts permanently installed screens whose moving images are not visible to the driver (625 ILCS 5/12-604.1). It also prohibits flashing lights except as expressly authorized and contains a broad catch-all against other unauthorized lighting (625 ILCS 5/12-212), and it restricts red, white, and blue flashing lights associated with emergency vehicles (625 ILCS 5/12-215, 12-215.1). Panels should be permanently mounted on the sides and rear, display static or slowly changing content while moving, and never show red light toward the front. Because the catch-all provision is broad and mobile billboard operators\u2019 claims are not legal authority, legal review should precede purchase.'));
C.push(P('Third, **licensing.** Chicago treats food pantries that meet three criteria as charitable food dispensing establishments exempt from the retail food license (City of Chicago 2026d), but a truck that dispenses prepared food falls under the city\u2019s mobile food rules, which require a commissary, a GPS device, daily servicing, insurance of at least $350,000, and no time-temperature-sensitive food stored in a parked vehicle (CDPH 2019; City of Chicago 2026a). Whether a free distribution vehicle needs a mobile food dispenser license is unresolved and should be settled in writing with the Chicago Department of Public Health. Off-premise advertising signs always require a permit under the municipal code (City of Chicago 2026b), although it is not clear that signs on a moving vehicle are covered.'));
C.push(P('Fourth, **dignity.** A brightly lit sponsor screen parked at a distribution stop can turn people receiving help into an audience for someone else\u2019s brand. The original proposal\u2019s idea of real-time "impact" counters makes this risk concrete. The program should adopt a written sponsor policy: no images of recipients without informed consent, displays dimmed or switched to neutral messaging during distribution, and impact reporting confined to aggregate figures published online. This protects recipients and, not incidentally, protects sponsors from reputational risk.'));

// ---------- 7 CBA ----------
C.push(H1('7. Cost-Benefit Analysis'));
C.push(H2('7.1 Benefits and Costs'));
C.push(P('The program modeled here is a fleet of one truck in year 1, two in year 2, and three from year 3 through year 10, a cap that reflects the realistic supply of partner restaurants in the central city. Each truck delivers 93,600 meals a year. Benefits have three parts. The largest is the value of food consumed: after assuming that 15 percent of rescued meals go uneaten, the midpoint of EPA\u2019s range, each consumed meal is valued at $2.07, the current Illinois SNAP benefit per meal, with sensitivity values of $0.90 (the Food Depository\u2019s cash cost per meal) and $3.70 (the national average meal cost). The second is the value of advertising services to sponsors, measured by the price paid. The third is net greenhouse gas reductions, valued at the $262 per ton of carbon dioxide equivalent implied by ReFED\u2019s 2024 estimates (ReFED 2026a). In 2025 the federal government disbanded the interagency working group on the social cost of greenhouse gases and withdrew its estimates (Executive Order 14154), so a zero value is tested as a sensitivity. Costs include all cash operating costs, capital purchases and replacement at the end of a seven-year life, and the opportunity cost of volunteer time. Benefits and costs of the Stage 1 status quo are subtracted.'));
C.push(H2('7.2 Results'));
C.push(P('The program passes a conventional cost-benefit test, though not by a wide margin (Table 5). At a 3 percent real discount rate, the present value of benefits is $4.65 million and of costs $4.20 million, for a net present value of $0.45 million and a benefit-cost ratio of 1.11. The result is stable across discount rates because most costs and benefits occur in the same years. Net annual flows are negative in the first three years as trucks are purchased and positive thereafter (Appendix C).'));
C.push(Cap('Table 5. Ten-year societal cost-benefit results, three-truck program'));
C.push(Tbl(['Discount rate (real)', 'PV benefits', 'PV costs', 'Net present value', 'Benefit-cost ratio'], [
  ...[['1.6% (A-94 ten-year rate)', MODEL['npv_0.016']], ['3% (A-4)', MODEL['npv_0.03']], ['7% (A-4)', MODEL['npv_0.07']]].map(([label, x]) =>
    [label, fmtM(x.pv_benefits), fmtM(x.pv_costs), fmtM(x.npv), x.bcr.toFixed(2)]),
], [2700, 1600, 1600, 1700, 1700]));
C.push(Note('Source: Author\u2019s model. Values are incremental to the Stage 1 status quo in constant 2026 dollars and include a conservative terminal value of the fleet in year 10 (Appendix A).'));
C.push(P('The climate result deserves emphasis because it runs against the usual framing of food recovery. Each truck diverts about 56 short tons of food a year from landfill, avoiding roughly 28 tons of carbon dioxide equivalent at WARM\u2019s landfill factor. But the truck and its refrigeration unit burn about 2,650 gallons of diesel, emitting about 27 tons. Under conservative accounting the net climate benefit is close to zero. It becomes meaningful only if rescued meals displace food that would otherwise have been produced; if half do, net avoided emissions rise to about 70 tons per truck per year. This is consistent with EPA\u2019s caution that donation carries energy costs (US EPA 2023a), and it implies that the program\u2019s case rests on feeding people, not on climate. It also points to a cheap improvement: a shorter, denser route or an electric van would make the climate benefit robust.'));
C.push(...Fig('fig5_cpm.png', 6.3, 2.71, 'Figure 5. Cost per meal: The Bridge Pantry compared with public and market benchmarks', 'Sources: Author\u2019s model; GCFD (2025a); USDA FNA (2026a); Feeding America (2026a).'));
C.push(H2('7.3 Sensitivity and Uncertainty'));
C.push(P('Figure 6 shows how the net present value responds to each assumption. Two parameters dominate: the value assigned to a meal and the number of meals each truck delivers per night. At 180 meals a night, the program loses money for society unless sponsorship is high; at 450, it is strongly positive under any sponsorship assumption. The share of food that goes uneaten and the sponsorship price are next in importance. Fuel prices, insurance, capital costs, and the discount rate matter much less. The practical implication is that the program\u2019s success will be determined by partner supply and distribution reliability, not by financial engineering.'));
C.push(...Fig('fig6_tornado.png', 6.3, 3.4, 'Figure 6. Sensitivity of program net present value to key assumptions', 'Source: Author\u2019s model. Each bar shows the NPV when one parameter moves from its low to its high value with all others at base values.'));
C.push(P('The Monte Carlo simulation draws all major parameters at once from triangular distributions (Figure 7). The median NPV is $0.93 million, the mean benefit-cost ratio is 1.24, and the program has a positive NPV in 78 percent of draws, with a 90 percent interval from \u2212$0.86 million to $3.39 million. A further way to read the result is the break-even meal value: the program\u2019s NPV is positive as long as each consumed meal is worth at least $1.79 to its recipient, which is below the $2.07 per meal that SNAP provides and less than half the $3.70 average cost of a meal. Because health and public-service savings are excluded, these results are conservative; even small reductions in emergency room use among frequent users would raise the benefit-cost ratio.'));
C.push(...Fig('fig7_mc.png', 6.3, 2.5, 'Figure 7. Monte Carlo distribution of program net present value (10,000 draws)', 'Source: Author\u2019s model.'));

// ---------- 8 Policy ----------
C.push(H1('8. Stage 3: An Illinois Edible-Food Recovery Framework'));
C.push(H2('8.1 Lessons from Other States'));
C.push(P('No U.S. state requires restaurants of every size to donate food. California and New York are the leading examples of states that require some generators to donate edible food and recycle the remainder (ReFED 2026c). California\u2019s SB 1383 requires covered commercial generators, including large grocery stores, food distributors, restaurants with 250 or more seats or more than 5,000 square feet, hotels with on-site food facilities and 200 or more rooms, large venues and events, and large health facilities, to donate edible food to food recovery organizations, and it set a statewide goal of recovering at least 20 percent of currently disposed edible food for people to eat by 2025 (ReFED 2026e). California pairs the mandate with incentives, including a tax credit worth 15 percent of the wholesale value of food donated to food banks by non-retail producers and processors and a credit for half the cost of transporting donated crops, and ReFED rates its liability protection as strong because state law also protects wholesome past-date food (ReFED 2026e). Massachusetts and Vermont took a different route, requiring organic waste to be separated and recycled rather than mandating donation: Massachusetts applies the rule to any non-residential entity producing more than half a ton of organic waste a week, and Vermont applies it to every person, business, and government entity in the state (ReFED 2026d; ReFED 2026g). The two states also illustrate how much guidance matters. Massachusetts\u2019s food safety guidance for donors is rated weak, while Vermont has issued state guidance on storage, labeling, and tracking for donations and school share tables; unclear safety rules are a leading reason businesses decline to donate (ReFED 2026d; ReFED 2026g).'));
C.push(Cap('Table 6. Policy comparison: Illinois and comparator states'));
C.push(Tbl(['Policy element', 'California', 'Vermont', 'Massachusetts', 'Illinois (current)', 'Proposed Illinois'], [
  ['Edible food donation required', 'Yes, covered generators', 'No', 'No', 'No', 'Yes, tiered (2028, 2030)'],
  ['Organic waste recycling requirement', 'Yes (strong)', 'Yes, all generators (strong)', '>\u00bd ton/week (moderate)', 'No', 'Phase 2, after donation tiers'],
  ['Donor liability beyond federal floor', 'Past-date food (strong)', 'Federal floor (moderate)', 'Past-date food (strong)', 'State act, 745 ILCS 50', 'Align with FDIA'],
  ['State law protects cost-recovery pricing', '\u2014', '\u2014', '\u2014', 'No; \u00a7 2.07 excludes sellers', 'Yes'],
  ['State tax incentives for donation', 'Food bank and transport credits', 'None', 'Farm credit, $5,000 cap', '\u2014', 'Transport credit'],
  ['Donor food safety guidance', 'Multiple guides (moderate)', 'State guidance (moderate)', 'Weak', '\u2014', 'IDPH prepared-food guidance'],
], [2100, 1450, 1400, 1400, 1450, 1500], { leftAll: true }));
C.push(Note('Sources: ReFED (2026c; 2026d; 2026e; 2026g); 745 ILCS 50; Food Donation Improvement Act of 2022. Ratings in parentheses are ReFED Policy Finder ratings. "\u2014" indicates the item is outside the scope of the sources reviewed.'));
C.push(H2('8.2 The Proposal'));
C.push(P('The paper proposes an Illinois Edible Food Recovery Act built in four parts. **Tier 1 (2028)** would require large generators with predictable surplus, including caterers above a revenue threshold, hotels with 200 or more rooms and on-site food service, stadiums and convention venues, hospitals, and colleges and universities with dining operations above a meal-count threshold, to maintain written agreements with food recovery organizations and report donated volumes annually. **Tier 2 (2030)** would extend the requirement to large restaurants using California\u2019s seat or floor-area thresholds and to restaurant chains above a location count. Small independent restaurants would be exempt but eligible for incentives. Organizations such as The Bridge Pantry would serve as registered logistics partners, and the state would publish a public registry matching generators with recovery organizations.'));
C.push(P('**Liability alignment.** Illinois law already protects restaurants and prepared food, but its definition of a protected nonprofit excludes organizations that sell donated food (745 ILCS 50/2.07). Federal law since 2023 protects nonprofits that charge a Good Samaritan reduced price no greater than the cost of handling and distribution, and protects restaurants and higher education institutions that donate directly to individuals (Food Donation Improvement Act of 2022; ReFED 2026f). Illinois should amend its act to match, which would allow social-enterprise caf\u00e9s and community grocery models to recover their handling costs without losing state protection.'));
C.push(P('**Safety guidance and infrastructure.** The Illinois Department of Public Health should issue plain-language guidance on donating time-temperature-controlled prepared food, including approved cooling, labeling, and transport practices, so that donor uncertainty does not undo the mandate. A targeted tax credit for the cost of transporting donated prepared food, modeled on California\u2019s credit for transporting donated crops (ReFED 2026e), would reward restaurants and caterers that deliver rather than wait for pickup. The state should also fund refrigerated vehicles, cold storage, and data systems for recovery organizations through competitive grants, which would position Illinois to use federal funds if pending legislation such as the Zero Food Waste Act of 2025 passes (ReFED 2026c). Illinois\u2019s senior senator is a lead sponsor of the bipartisan Reduce Food Loss and Waste Act of 2025, which offers a natural federal-state partnership (ReFED 2026c). Finally, the state or city should clarify how the Vehicle Code\u2019s lighting provisions apply to permanently mounted, non-flashing sponsor displays on nonprofit vehicles.'));
C.push(H2('8.3 Statewide Impact'));
C.push(P('Table 7 estimates what the framework could achieve. If the share of Illinois foodservice surplus that is recovered rose from 0.24 percent to 5 percent, a level well within the overproduction share of restaurant surplus, the state would recover about 22,700 tons a year, enough for 37.9 million meals or about 729,000 meals a week. That is roughly seven times the original proposal\u2019s statewide target of 100,000 meals a week. It would also require logistics capacity equivalent to about 400 Bridge Pantry trucks, which is why the framework must be paired with infrastructure funding and with the Food Depository and other food banks rather than relying on small organizations alone.'));
C.push(Cap('Table 7. Statewide recovery scenarios for Illinois foodservice surplus'));
C.push(Tbl(['Share of foodservice surplus recovered', 'Tons per year', 'Meals per year', 'Meals per week', 'Truck-equivalents', 'Landfill CO\u2082e avoided (t)'], [
  ['Current (' + (100 * MODEL.fs_donated / MODEL.fs_surplus).toFixed(2) + '%)', fmtInt(MODEL.fs_donated), fmtMeals(MODEL.fs_donated * 2000 / 1.2),
    fmtInt(round100(MODEL.fs_donated * 2000 / 1.2 / 52)), fmtInt(MODEL.fs_donated * 2000 / 1.2 / T.meals), fmtInt(MODEL.fs_donated * 0.5)],
  ...MODEL.statewide.map(x => [Math.round(x.share * 100) + '%', fmtInt(x.tons), fmtMeals(x.meals), fmtInt(round100(x.meals_week)), fmtInt(x.trucks), fmtInt(x.ghg_landfill)]),
], [2500, 1250, 1350, 1350, 1350, 1500]));
C.push(Note('Source: Author\u2019s calculations from ReFED (2026a), 2024 Illinois foodservice surplus of ' + fmtInt(MODEL.fs_surplus) + ' tons. Meals at 1.2 lb each; truck-equivalents at ' + fmtInt(T.meals) + ' meals per truck per year; emissions use the WARM landfill factor only.'));

// ---------- 9 SNAP ----------
C.push(H1('9. Complementing, Not Replacing, SNAP'));
C.push(P('The original proposal suggested redirecting part of SNAP toward housing and replacing the food it buys with recovered surplus. That idea should be set aside, for legal, empirical, and ethical reasons. SNAP benefits are restricted by federal law to food purchases, so redirecting them to rent would require an act of Congress. Recovered food is unpredictable in quantity and type, so it cannot provide the steady, chosen diet that SNAP provides. And the 2025 federal changes have already reduced Illinois SNAP enrollment by about a fifth; presenting food recovery as a substitute would invite further cuts precisely when need is rising. The literature on charitable food is clear that it works best as a supplement to public entitlements (Poppendieck 1998; McIntyre et al. 2016).'));
C.push(P('Food recovery can complement SNAP in three concrete ways. First, it is cheap per meal. At $1.33 to $1.44 per meal in cash costs, recovered prepared meals deliver food value comparable to SNAP\u2019s $2.07 per meal at a lower public cost, which justifies state co-investment in recovery infrastructure as a hunger program in its own right. Second, it reaches people SNAP serves poorly. Unsheltered adults often cannot store or cook groceries, which is why Illinois allows homeless SNAP households to buy hot meals at participating restaurants through the Restaurant Meals Program. Participating restaurants give a 10 percent discount and may not charge tax (IDHS 2026), and the state expanded the program to 50 restaurants in 16 counties in October 2026 (WSIU 2026). Recruiting partner restaurants such as Nando\u2019s into the program in the Loop and on the South Side would give the same people both free recovered meals and the ability to spend their own benefits on hot food. Third, a recovery program is a SNAP enrollment channel: TEFAP rules already require distribution sites to accept self-declared homelessness and encourage referral to SNAP (IDHS 2024), and outreach workers on a nightly route can help people keep benefits under new work requirements.'));
C.push(P('The research questions in the original outline can therefore be answered directly. The cost per meal delivered is about $1.44 in the base case and $1.33 in the lean case. Each dollar of program cash spending delivers about 0.69 to 0.75 meals and, after allowing for food that goes uneaten, about $1.22 to $1.33 of food value at SNAP prices. The defensible claim is not that food recovery saves SNAP money, but that it stretches the public dollar furthest for the people SNAP reaches least.'));

// ---------- 10 Stage 4 ----------
C.push(H1('10. Stage 4: Reintegration Centers and Social-Enterprise Caf\u00e9s'));
C.push(P('Once Stage 2 operations are stable, the nightly route becomes something more valuable than a meal service: a trusted, predictable point of contact with people who are otherwise hard to reach. Stage 4 uses that trust to build a pathway out of homelessness (Figure 8). The proposed reintegration center would be open to everyone, with no religious requirement and no discrimination, offering showers, laundry, mail and identification services, benefits enrollment, counseling, and case management. Church-based outreach centers have long provided these services in Chicago; the design here borrows their low-barrier hospitality while keeping any religious activity separate in time and place and voluntary, as federal rules require of organizations receiving USDA foods (7 C.F.R. \u00a7 16.4(b); IDHS 2024).'));
C.push(...Fig('fig8_pathway.png', 6.4, 1.6, 'Figure 8. The Stage 4 pathway from a nightly meal to stable housing', 'Source: Author.'));
C.push(P('The attached caf\u00e9 would serve as a paid training site, staffed partly by program participants, with placement in partner restaurants as the goal. The evidence supports the model but tempers expectations about its finances. Social enterprise employment improves job retention and housing stability (Rotz, Maxwell, and Dunn 2015), but Inspiration Kitchens, the closest Chicago analogue, covers only about 10 percent of its training program\u2019s costs from restaurant revenue (Inspiration Corporation 2026). The caf\u00e9 should be budgeted as a subsidized workforce program, with partner restaurants such as Nando\u2019s contributing training slots and hiring commitments as part of their CSR partnership. Legally, rescued food can be served free in the caf\u00e9 or sold only at a price that covers handling costs, which the federal law permits but current Illinois law does not protect; until the state act is amended, the caf\u00e9\u2019s sales menu should use purchased ingredients (745 ILCS 50/2.07; Food Donation Improvement Act of 2022). Finally, the center should not try to become a housing provider. Its job is to connect people to Chicago\u2019s coordinated entry system, which housed 13 percent more people in 2025 than in 2024 (DFSS 2026), and to Housing First programs with the strongest evidence of keeping people housed (Tsemberis, Gulcur, and Nakae 2004).'));

// ---------- 11 Metrics ----------
C.push(H1('11. Impact Measurement Framework'));
C.push(P('Credibility with sponsors, funders, and legislators depends on measurement. Table 8 sets out a compact dashboard. Operational metrics should be recorded nightly from the pickup log and route data the truck\u2019s GPS already requires, financial metrics monthly, and outcome metrics quarterly through case management records. The most important discipline is weighing food at pickup, which replaces the conversion assumptions used in this paper with measured data within the first month of Stage 2.'));
C.push(Cap('Table 8. Impact metrics dashboard'));
C.push(Tbl(['Domain', 'Metric', 'Source', 'Year 1 target (lean pilot)'], [
  ['Environmental', 'Pounds rescued; landfill diversion', 'Scale at pickup', '82,000 lb'],
  ['Environmental', 'Net CO\u2082e avoided (incl. vehicle)', 'WARM factors, fuel log', 'Positive net'],
  ['Social', 'Meals served; unique people reached', 'Route tallies', '68,000 meals'],
  ['Social', 'Referrals to SNAP, ID, coordinated entry', 'Case log', '150 referrals'],
  ['Social', 'Training enrollments and job placements', 'Caf\u00e9 records (Stage 4)', 'Baseline year'],
  ['Financial', 'Cash cost per meal', 'Accounts', '\u2264 $1.40'],
  ['Financial', 'Operating coverage (contributed \u00f7 cost)', 'Accounts', '\u2265 80%'],
  ['Compliance', 'Pickups with temperature record', 'Pickup log', '100%'],
], [1600, 3300, 2200, 2200], { leftAll: true }));

// ---------- 12 Implementation ----------
C.push(H1('12. Implementation Roadmap and Risks'));
C.push(P('Table 9 sets out a realistic sequence. It is slower than the original 18-month roadmap because procurement, licensing, and sponsor agreements take time, and because each stage should begin only when the previous one is producing reliable data.'));
C.push(Cap('Table 9. Implementation roadmap'));
C.push(Tbl(['Period', 'Milestone'], [
  ['Q4 2026', 'Pickup log and food-weighing protocol; formal agreements with informal sources; written guidance request to CDPH; legal review of display design and sponsorship terms'],
  ['Q1\u2013Q2 2027', 'Lean pilot vehicle; Nando\u2019s anchor sponsorship (acknowledgment-only); add 4\u20136 partner restaurants to reach nightly pickups; partnership discussions with the Food Depository'],
  ['Q3 2027\u2013Q2 2028', 'Full-size refrigerated truck if pilot exceeds 220 meals/night and 80% operating coverage; second sponsor slot; begin Restaurant Meals Program recruitment'],
  ['2028', 'Policy advocacy for the Illinois Edible Food Recovery Act and liability alignment; second truck'],
  ['2029\u20132030', 'Reintegration center and training caf\u00e9 in partnership with an existing service provider; third truck'],
], [1800, 7500], { leftAll: true }));
C.push(P('The main risks follow from the analysis. Supply risk, the possibility that partner surplus shrinks as restaurants reduce waste, is the most serious and is partly a sign of success; it argues for a diversified partner base and for measuring source reduction as an outcome. Labor risk, dependence on a few committed volunteers, is addressed by the paid coordinator role. Financial risk is addressed by blended revenue and by not purchasing the full-size truck until the pilot proves demand. Legal risk is addressed by acknowledgment-only sponsorship, permanent non-flashing displays, written health department guidance, and documented food handling. Reputational risk is addressed by the dignity policy in Section 6.3. Finally, the program depends on access to the places unhoused people stay. Across the country, cities have restricted sharing food with people experiencing homelessness through permit and fee requirements for using public property, mandates to comply with food-safety codes, and neighborhood pressure; at least 21 cities adopted such restrictions between January 2013 and late 2014 alone (NCH 2014), and the National Homelessness Law Center tracks these measures alongside broader criminalization laws in 187 cities (NCH and NLCHP 2010; NHLC 2026). Street distribution is exactly the activity these rules target, so the program should coordinate stops with city outreach teams, document its food-safety practices, which answers the most common regulatory objection, and treat any future local restriction as a policy issue to contest through advocacy rather than an operational surprise.'));

// ---------- 13 Conclusion ----------
C.push(H1('13. Conclusion'));
C.push(P('The Bridge Pantry demonstrates that the distance between surplus and need in Chicago can be measured in blocks and minutes. The question this paper asked was whether that insight can become a durable institution. The answer is a qualified yes. A sponsored, refrigerated food recovery truck can deliver meals at a lower cost than either public or market benchmarks, and a three-truck program passes a conservative cost-benefit test with a high probability of a positive return. But advertising alone will not sustain it, its climate benefits are smaller than often claimed, and its success depends far more on reliable restaurant partners and distribution than on financial structure.'));
C.push(P('Those conclusions point to a clear strategy. Start lean, measure everything, and treat sponsorship as one revenue stream among several. Partner with the Food Depository rather than duplicating it. Ask the state to do what only the state can: require large generators to offer their edible surplus, align Illinois liability law with federal law, issue clear food safety guidance, and fund the cold chain. Keep SNAP whole and use recovery to reach the people SNAP reaches least. And build toward reintegration services that turn a nightly meal into a path to work and housing. If the model succeeds in Chicago, it offers a replicable template for other Illinois cities, and Illinois foodservice businesses, which today donate less than one percent of their surplus, have ample room to supply it.'));

// ---------- References ----------
C.push(PB());
C.push(H1('References'));
const refs = [
  'Alexander, Catherine, and Chris Smaje. 2008. \u201cSurplus Retail Food Redistribution: An Analysis of a Third Sector Model.\u201d *Resources, Conservation and Recycling* 52 (11): 1290\u201398.',
  'Bacon, Christopher M., and Gregory A. Baker. 2017. \u201cThe Rise of Food Banks and the Challenge of Matching Food Assistance with Potential Need: Towards a Spatially Specific, Rapid Assessment Approach.\u201d *Agriculture and Human Values* 34 (4): 899\u2013919.',
  'Boost Outdoor Media. 2026. \u201cChicago Billboard and Outdoor Advertising Guide: Costs, Locations and LED Trucks.\u201d September 30, 2026.',
  'Can\u2019t Miss US. 2025. \u201cBillboard Truck Advertising: How Much Does a Digital Billboard Truck Cost to Own.\u201d September 26, 2025.',
  'CDPH (Chicago Department of Public Health). 2019. *Mobile Food Vendors and Shared Kitchens Rules*. Effective January 1, 2019. Chicago: CDPH.',
  'City of Chicago. 2026a. \u201cFood Trucks.\u201d Department of Business Affairs and Consumer Protection, Business Licensing. Accessed October 14, 2026.',
  'City of Chicago. 2026b. \u201cSign Permits.\u201d Department of Buildings. Accessed October 14, 2026.',
  'City of Chicago. 2026c. \u201cMinimum Wage.\u201d Department of Business Affairs and Consumer Protection, Office of Labor Standards. Accessed October 14, 2026.',
  'City of Chicago. 2026d. \u201cGuidance on Opening Food Pantries.\u201d Advancing Food Equity in Chicago. Accessed October 14, 2026.',
  'Culhane, Dennis P., Stephen Metraux, and Trevor Hadley. 2002. \u201cPublic Service Reductions Associated with Placement of Homeless Persons with Severe Mental Illness in Supportive Housing.\u201d *Housing Policy Debate* 13 (1): 107\u201363.',
  'Cunnyngham, Karen. 2025. *Reaching Those in Need: Estimates of State SNAP Participation Rates in 2022*. Report prepared by Mathematica for the U.S. Department of Agriculture, Food and Nutrition Service. February 2025.',
  'DFSS (Chicago Department of Family and Support Services). 2026. *2026 Snapshot of Homelessness in Chicago*. Point-in-Time Count one-page summary. Chicago: DFSS.',
  'Executive Order 14154. 2025. \u201cUnleashing American Energy.\u201d January 20, 2025.',
  'Feeding America. 2026a. *Map the Meal Gap 2026*. Chicago: Feeding America. Released July 28, 2026.',
  'Feeding America. 2026b. *Map the Meal Gap 2026: Technical Brief*. Chicago: Feeding America.',
  'Galli, Francesca, Alessio Cavicchi, and Gianluca Brunori. 2019. \u201cFood Waste Reduction and Food Poverty Alleviation: A System Dynamics Conceptual Model.\u201d *Agriculture and Human Values* 36 (2): 289\u2013300.',
  'GCFD (Greater Chicago Food Depository). 2023a. *2023 Annual Report* and *Supporters 2022\u20132023*. Chicago: GCFD.',
  'GCFD. 2023b. Form 990, Return of Organization Exempt from Income Tax, Fiscal Year Ended June 30, 2023. Public disclosure copy.',
  'GCFD. 2024a. *2024 Annual Report* and *Supporters 2023\u20132024*. Chicago: GCFD.',
  'GCFD. 2024b. Form 990, Return of Organization Exempt from Income Tax, Fiscal Year Ended June 30, 2024. Public disclosure copy.',
  'GCFD. 2025a. *2025 Annual Report*, audited financial statements, and *Supporters 2024\u20132025*. Chicago: GCFD.',
  'GCFD. 2025b. Form 990, Return of Organization Exempt from Income Tax, Fiscal Year Ended June 30, 2025. Filed December 19, 2025.',
  'GCFD. 2025c. Form 990-T, Exempt Organization Business Income Tax Return, Fiscal Year Ended June 30, 2025. Filed December 19, 2025.',
  'GCFD. 2026a. *2026 Hunger in Chicago Report*. Survey conducted with NORC at the University of Chicago. Chicago: GCFD.',
  'Gundersen, Craig, and James P. Ziliak. 2015. \u201cFood Insecurity and Health Outcomes.\u201d *Health Affairs* 34 (11): 1830\u201339.',
  'IDHS (Illinois Department of Human Services). 2024. *The Emergency Food Assistance Program Distribution Site Manual for Food Pantries, Soup Kitchens, and Homeless Shelters, State Fiscal Year 2025*. Updated May 3, 2024. Springfield: IDHS.',
  'IDHS. 2026. \u201cSNAP Restaurant Meals Program\u201d and \u201cRestaurant Meals: SNAP Customers FAQ.\u201d Accessed October 14, 2026.',
  'IDPH (Illinois Department of Public Health). n.d. \u201cFood Banks: Illinois Food Handling Regulation Act, 410 ILCS 625/3.2.\u201d Guidance flyer.',
  'Inspiration Corporation. 2026. \u201cWhat We Do.\u201d Accessed October 14, 2026. https://www.inspirationcorp.org/what-we-do.',
  'Lee, Deishin, Erkut S\u00f6nmez, Miguel I. G\u00f3mez, and Xiaoli Fan. 2017. \u201cCombining Two Wrongs to Make Two Rights: Mitigating Food Insecurity and Food Waste through Gleaning Operations.\u201d *Food Policy* 68: 40\u201352.',
  'Lohnes, Joshua D. 2021. \u201cRegulating Surplus: Charity and the Legal Geographies of Food Waste Enclosure.\u201d *Agriculture and Human Values* 38 (2): 351\u201363.',
  'McIntyre, Lynn, Danielle Tougas, Krista Rondeau, and Catherine L. Mah. 2016. \u201c\u2018In\u2019-sights about Food Banks from a Critical Interpretive Synthesis of the Academic Literature.\u201d *Agriculture and Human Values* 33 (4): 843\u201359.',
  'NCH (National Coalition for the Homeless). 2014. *Share No More: The Criminalization of Efforts to Feed People in Need*. Washington, DC: NCH. October 2014.',
  'NCH (National Coalition for the Homeless) and NLCHP (National Law Center on Homelessness and Poverty). 2010. *A Place at the Table: Prohibitions on Sharing Food with People Experiencing Homelessness*. Washington, DC.',
  'Nelsen, M., Virginia Gray, and Samuel Paler-Ponce. 2026. *2026 Estimate of People Experiencing Homelessness in Chicago: Reporting Data 2015\u20132025*. Chicago: Chicago Coalition to End Homelessness.',
  'NHLC (National Homelessness Law Center, formerly the National Law Center on Homelessness and Poverty). 2026. \u201cCriminalization.\u201d Reports and resources, including *Housing Not Handcuffs 2019* and the *2021 State Law Supplement*. Accessed October 14, 2026.',
  'OMB (Office of Management and Budget). 2003. *Circular A-4: Regulatory Analysis*. September 17, 2003. Washington, DC: Executive Office of the President.',
  'OMB. 2025. *Recission and Reinstatement of Circular A-4*. Memorandum M-25-15, February 12, 2025.',
  'OMB. 2026. *Circular A-94, Appendix C: Discount Rates for Cost-Effectiveness, Lease Purchase, and Related Analyses*. Revised March 6, 2026.',
  'Poppendieck, Janet. 1998. *Sweet Charity? Emergency Food and the End of Entitlement*. New York: Viking.',
  'Porter, Michael E., and Mark R. Kramer. 2011. \u201cCreating Shared Value.\u201d *Harvard Business Review* 89 (1\u20132): 62\u201377.',
  'REDF. 2015. \u201cNew Study Reveals Society\u2019s Return on Investment More than Doubles with Employment through Social Enterprise.\u201d Press release, February 5, 2015.',
  'ReFED. 2026a. *Insights Engine: U.S. and State Food Surplus Data* (Food Surplus Summary and Cause Summary files). Data last updated April 3, 2026.',
  'ReFED. 2026b. \u201cFoodservice Methodology.\u201d *Insights Engine Documentation*. Accessed October 14, 2026.',
  'ReFED. 2026c. \u201cCurrent Federal Policy.\u201d *U.S. Food Waste Policy Finder*. Accessed October 14, 2026.',
  'ReFED. 2026d. \u201cMassachusetts.\u201d *U.S. Food Waste Policy Finder*. Accessed October 14, 2026.',
  'ReFED. 2026e. \u201cCalifornia.\u201d *U.S. Food Waste Policy Finder*. Accessed October 14, 2026.',
  'ReFED. 2026f. \u201cFederal Liability Protection.\u201d *U.S. Food Waste Policy Finder*. Accessed October 14, 2026.',
  'ReFED. 2026g. \u201cVermont.\u201d *U.S. Food Waste Policy Finder*. Accessed October 14, 2026.',
  'Richard, Molly K., Julie Dworkin, Katherine Grace Rule, Suniya Farooqui, Zachary Glendening, and Sam Paler-Ponce. 2022. \u201cQuantifying Doubled-Up Homelessness: Presenting a New Measure Using U.S. Census Microdata.\u201d *Housing Policy Debate*. https://doi.org/10.1080/10511482.2021.1981976.',
  'Rotz, Dana, Nan Maxwell, and Adam Dunn. 2015. *Economic Self-Sufficiency and Life Stability One Year after Starting a Social Enterprise Job*. Oakland, CA: Mathematica Policy Research.',
  'Tarasuk, Valerie, and Joan M. Eakin. 2005. \u201cFood Assistance through \u2018Surplus\u2019 Food: Insights from an Ethnographic Study of Food Bank Work.\u201d *Agriculture and Human Values* 22 (2): 177\u201386.',
  'The Bridge Pantry Chicago. 2025. \u201cTurning Food Waste into Community Impact.\u201d Executive presentation for Nando\u2019s. Unpublished.',
  'The Bridge Pantry Chicago. 2026. \u201cHarvest Program: Pickup Route for Nando\u2019s PERi-PERi Chicago Locations.\u201d Internal route sheet. Unpublished.',
  'Tsemberis, Sam, Leyla Gulcur, and Maria Nakae. 2004. \u201cHousing First, Consumer Choice, and Harm Reduction for Homeless Individuals with a Dual Diagnosis.\u201d *American Journal of Public Health* 94 (4): 651\u201356.',
  'US EIA (U.S. Energy Information Administration). 2026a. *U.S. On-Highway Diesel Fuel Prices* (psw18vwall). Weekly and monthly series, released October 14, 2026.',
  'US EIA. 2026b. *Gasoline and Diesel Fuel Update*. Weekly retail prices for October 12, 2026, released October 14, 2026.',
  'US EPA (U.S. Environmental Protection Agency). 2023a. *From Field to Bin: The Environmental Impacts of U.S. Food Waste Management Pathways (Part 2)*. Washington, DC: EPA Office of Research and Development. October 2023.',
  'US EPA. 2023b. *Documentation for Greenhouse Gas Emission and Energy Factors Used in the Waste Reduction Model (WARM), Version 16: Organic Materials Chapters*. December 2023.',
  'US EPA. 2023c. *Waste Reduction Model (WARM) Version 16 User\u2019s Guide*. December 2023.',
  'USDA FNA (U.S. Department of Agriculture, Food and Nutrition Administration). 2026a. *SNAP Data Tables*: national annual summary and state-level persons, households, and benefits. Data as of September 11, 2026.',
  'USDA FNA. 2026b. *SNAP State Level Participation and Benefits, FY1989\u2013FY2026* (National Data Bank public files). Released September 10, 2026.',
  'Varadarajan, P. Rajan, and Anil Menon. 1988. \u201cCause-Related Marketing: A Coalignment of Marketing Strategy and Corporate Philanthropy.\u201d *Journal of Marketing* 52 (3): 58\u201374.',
  'WSIU. 2026. \u201cIllinois Expands Restaurant Meals Program for Eligible SNAP Customers.\u201d October 8, 2026.',
];
refs.forEach(r => C.push(Ref(r)));
C.push(H2('Legal Sources'));
[
  '7 C.F.R. \u00a7 16.4 (Equal Opportunity for Religious Organizations).',
  'Cal. Code Regs. tit. 14, \u00a7\u00a7 18981.1\u201318998.4 (SB 1383 short-lived climate pollutants: organic waste reduction and edible food recovery).',
  '26 U.S.C. \u00a7 513(i) and 26 C.F.R. \u00a7 1.513-4 (qualified sponsorship payments).',
  '42 U.S.C. \u00a7 1791 (Bill Emerson Good Samaritan Food Donation Act), as amended by the Food Donation Improvement Act of 2022, Pub. L. No. 117-362 (2023).',
  '410 ILCS 625/3.2 (Illinois Food Handling Regulation Enforcement Act, food banks).',
  '625 ILCS 5/12-101, 12-212, 12-215, 12-215.1, 12-604.1 (Illinois Vehicle Code).',
  '745 ILCS 50 (Illinois Good Samaritan Food Donor Act).',
  'Vt. Stat. Ann. tit. 10, \u00a7\u00a7 6602, 6605k (Universal Recycling Law).',
  'Municipal Code of Chicago, ch. 4-8 (food establishments and mobile food vendors) and ch. 13-20, art. XIII (signs).',
].forEach(r => C.push(Ref(r)));

// ---------- Appendices ----------
C.push(PB());
C.push(H1('Appendix A. Model Assumptions and Sources'));
C.push(Cap('Table A1. Parameters used in the cost-benefit and Monte Carlo analysis'));
C.push(Tbl(['Parameter', 'Base', 'Low', 'High', 'Source or basis'], [
  ['Operating nights per year', '312', '\u2014', '\u2014', 'Six nights per week'],
  ['Meals per truck-night', '300', '180', '450', 'Partner supply: 8 Nando\u2019s + informal + new partners'],
  ['Pounds per meal', '1.2', '\u2014', '\u2014', 'ReFED (2026a); GCFD reporting'],
  ['Meals not consumed', '15%', '30%', '5%', 'US EPA (2023a) range 4\u201340%'],
  ['Value per consumed meal', '$2.07', '$0.90', '$3.70', 'SNAP benefit of $189.09 \u00f7 (30.4 days \u00d7 3 meals), USDA FNA (2026a); GCFD (2025a); Feeding America (2026a)'],
  ['Diesel price ($/gal)', '$5.00', '$6.26', '$3.62', 'EIA (2026a): U.S. Jan.\u2013Sep. 2026 average (Midwest $4.94); Midwest, week of Oct. 12, 2026; Midwest 2025 average'],
  ['Driver wage ($/hr)', '$22.00', '$26.00', '$19.00', 'Above Chicago minimum of $17.05 (City of Chicago 2026c)'],
  ['Fringe benefit rate', '26.4%', '\u2014', '\u2014', 'GCFD (2025a) audited statements'],
  ['Insurance ($/yr)', '$8,000', '$12,000', '$6,000', 'Planning estimate; $350,000 city minimum'],
  ['Truck + LED + wrap', '$140,000', '$170,000', '$110,000', 'Appendix B; Can\u2019t Miss US (2025); Boost Outdoor Media (2026)'],
  ['Vehicle life and replacement', '7 years', '\u2014', '\u2014', 'GCFD (2025a) lease terms; replaced net of 15% resale value'],
  ['Terminal value of fleet (year 10)', '45% of price', '\u2014', '\u2014', 'Conservative; below straight-line book value'],
  ['Sponsorship ($/month)', '$4,000', '$1,500', '$8,000', 'Section 6.2'],
  ['CSR contributions ($/yr)', '$20,000', '$5,000', '$40,000', 'Planning estimate'],
  ['Landfill emissions factor', '0.50 t/ton', '\u2014', '\u2014', 'US EPA (2023b), WARM v16'],
  ['Share of meals displacing production', '0%', '0%', '50%', 'US EPA (2023a); poultry factor 2.45 t/ton'],
  ['Social cost of carbon ($/t)', '$262', '$0', '$262', 'ReFED (2026a); Executive Order 14154'],
  ['Volunteer time ($/hr)', '$17.05', '\u2014', '\u2014', 'City of Chicago (2026c)'],
  ['Discount rate (real)', '3%', '7%', '1.6%', 'OMB (2003; 2026)'],
], [2400, 1000, 900, 900, 4100], { leftAll: true }));
C.push(Note('Low and high values are those that lower and raise net present value, respectively. Monte Carlo draws use triangular distributions over these ranges, with the displacement share drawn uniformly between 0 and 50 percent.'));

C.push(H1('Appendix B. Stage 2 Capital and Start-Up Cost Estimates'));
C.push(PN('The figures below are planning estimates for budgeting and fundraising. They are derived from published vendor price ranges, Chicago fee schedules, and comparable nonprofit fleet costs, and should be replaced with at least two written vendor quotes before procurement.'));
C.push(Cap('Table B1. Planning estimates for Stage 2 vehicles and equipment'));
C.push(Tbl(['Item', 'Full-size base case', 'Lean pilot', 'Notes'], [
  ['Refrigerated box truck (16 ft, non-CDL)', '$95,000 (new)', '\u2014', 'Range $80,000\u2013$110,000'],
  ['Used refrigerated cargo van', '\u2014', '$55,000', 'Lower capacity, ~220 meals/night'],
  ['LED side and rear display system, installed', '$40,000', '$25,000', 'Permanently mounted; non-flashing content'],
  ['Sponsor wrap and branding', '$5,000', '$5,000', 'Acknowledgment-only design'],
  ['Hot-holding and cold-holding equipment', 'Included', 'Included', 'Insulated carriers, thermometers, data loggers'],
  ['Licensing and fire safety review (first cycle)', '$950', '$950', 'Dispenser license $700/2 yrs; fire permit $100; review $150'],
  ['**Total start-up cost**', '**$140,950**', '**$85,950**', ''],
], [3300, 1800, 1500, 2700], { leftAll: true }));
C.push(Note('Sources: Can\u2019t Miss US (2025); Boost Outdoor Media (2026); City of Chicago (2026a); GCFD (2025a) for comparable fleet lease costs of about $37,000 per vehicle per year (total lease expense divided by leased vehicles), which can be used to evaluate leasing rather than buying with the OMB (2026) lease-purchase discount rates.'));

C.push(H1('Appendix C. Ten-Year Program Cash Flows'));
C.push(Cap('Table C1. Societal benefits and costs by year, incremental to Stage 1 (3% real, undiscounted values shown)'));
const flows = MODEL.flows_3;
const orgf = MODEL.org_flows;
const fm = v => (v < 0 ? '\u2212$' : '$') + Math.abs(Math.round(v / 1000)).toLocaleString('en-US') + 'k';
C.push(Tbl(['Year', 'Trucks', 'Societal benefits', 'Societal costs', 'Net societal', 'Organization net cash'],
  flows.map((f, i) => [String(f.t), String(f.trucks), fm(f.benefits), fm(f.costs), fm(f.net), fm(orgf[i].net)]),
  [900, 900, 1900, 1900, 1800, 1900]));
C.push(Note('Years 8 to 10 include replacement of the trucks bought in years 1 to 3, net of their resale value; year 10 also includes the terminal value of the fleet. Organization net cash equals contributed revenue less cash operating costs and capital purchases, and shows the annual fundraising gap the blended-finance model must close.'));

// ---------- Contents page (static, page numbers from outputs/toc_pages.json) ----------
let pages = {};
try { pages = JSON.parse(fs.readFileSync(path.join(OUT, 'toc_pages.json'))); } catch (e) { /* first build: no page numbers yet */ }
const tocParas = HEADS.filter(h => h[1] !== 'Legal Sources').map(([lvl, t]) => new Paragraph({
  spacing: { after: lvl === 1 ? 60 : 30, line: 240 }, indent: { left: lvl === 1 ? 0 : 360 },
  tabStops: [{ type: TabStopType.RIGHT, position: 9360, leader: 'dot' }],
  children: [new TextRun({ text: t, font: FONT, size: lvl === 1 ? 22 : 21, bold: lvl === 1 }),
    new TextRun({ children: [new Tab()], font: FONT, size: 21 }),
    new TextRun({ text: String(pages[t] || ''), font: FONT, size: 21 })]
}));
C.splice(TOC_INDEX, 1, ...tocParas);
fs.writeFileSync(path.join(OUT, 'headings.json'), JSON.stringify(HEADS));

// ======================= DOCUMENT =======================
const doc = new Document({
  creator: 'Muhammad Jibrail', lastModifiedBy: 'Muhammad Jibrail',
  title: 'The Bridge Pantry: A Financially Sustainable Model for Reducing Food Waste and Homelessness in Chicago',
  styles: {
    default: { document: { run: { font: FONT, size: BODY_SIZE } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 28, bold: true, font: FONT, color: '1F3A5F' }, paragraph: { spacing: { before: 300, after: 140 }, outlineLevel: 0, keepNext: true } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 25, bold: true, font: FONT }, paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1, keepNext: true } },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }, titlePage: true },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'The Bridge Pantry: A Sustainable Food Recovery Model for Chicago', font: FONT, size: 18, italics: true, color: '666666' })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] }) },
    children: C
  }]
});
Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(path.join(OUT, 'Bridge_Pantry_Research_Paper.docx'), b);
  console.log('wrote outputs/Bridge_Pantry_Research_Paper.docx');
});
