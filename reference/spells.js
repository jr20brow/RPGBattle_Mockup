const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
    PageOrientation
} = require('docx');
const fs = require('fs');

// ─── COLOUR PALETTE ──────────────────────────────────────────────────────────
const COLS = {
    Fire:    { header: 'C0392B', light: 'F9EBEA' },
    Ice:     { header: '2471A3', light: 'EAF4FB' },
    Thunder: { header: 'B7950B', light: 'FEF9E7' },
    Wind:    { header: '1E8449', light: 'EAFAF1' },
    Earth:   { header: '784212', light: 'FAF0E6' },
    Grass:   { header: '1D6A39', light: 'E9F7EF' },
    Dark:    { header: '4A235A', light: 'F5EEF8' },
    Light:   { header: '9A7D0A', light: 'FDFBE7' },
    Anima:   { header: '1A5276', light: 'EAF2FB' },
};

// ─── SPELL DATA ───────────────────────────────────────────────────────────────
// Tiers: Basic (1T) | Powerful (1T) | AoE | Strong AoE ×2
// SP cost scale: Basic 6–8 | Powerful 12–16 | AoE 14–18 | Strong AoE ×2 20–26
// POW scale:     Basic 65  | Powerful 90–95 | AoE 70    | Strong AoE ×2 85 (×2)

const SPELLS = [
    {
        element: 'Fire',
        spells: [
            { tier: 'Basic — Single Target',       name: 'Fireball',        pow: 65,  cost: 6,  hits: 1, note: 'A basic Fire spell against a single target. Already in game.' },
            { tier: 'Powerful — Single Target',    name: 'Tower of Flame',  pow: 90,  cost: 14, hits: 1, note: '10% chance to inflict Burn on the target for 3 rounds. Already in game.' },
            { tier: 'AoE',                         name: 'Eruption',        pow: 70,  cost: 16, hits: 1, note: 'A pillar of fire erupts beneath all enemies.' },
            { tier: 'Strong AoE — Hits Twice',     name: 'Inferno',         pow: 85,  cost: 24, hits: 2, note: 'A roaring firestorm washes over all enemies twice. 15% chance of Burn per hit.' },
        ]
    },
    {
        element: 'Ice',
        spells: [
            { tier: 'Basic — Single Target',       name: 'Chill',           pow: 65,  cost: 6,  hits: 1, note: 'A basic Ice spell against a single target. Already in game.' },
            { tier: 'Powerful — Single Target',    name: 'Glacial Lance',   pow: 92,  cost: 13, hits: 1, note: 'A shard of concentrated ice pierces a single target. 20% chance to reduce target SPD by 1 for 2 rounds.' },
            { tier: 'AoE',                         name: 'Frost Nova',      pow: 70,  cost: 16, hits: 1, note: 'A burst of freezing air engulfs all enemies.' },
            { tier: 'Strong AoE — Hits Twice',     name: 'Blizzard',        pow: 85,  cost: 23, hits: 2, note: 'A sustained ice storm batters all enemies twice.' },
        ]
    },
    {
        element: 'Thunder',
        spells: [
            { tier: 'Basic — Single Target',       name: 'Thunderbolt',     pow: 65,  cost: 6,  hits: 1, note: 'A basic Thunder spell against a single target. Already in game.' },
            { tier: 'Powerful — Single Target',    name: 'Lightning Strike', pow: 95, cost: 14, hits: 1, note: 'A concentrated bolt strikes a single target. Ignores 50% of the target\'s M. Def.' },
            { tier: 'AoE',                         name: 'Static Field',    pow: 70,  cost: 15, hits: 1, note: 'An electric pulse radiates outward, hitting all enemies.' },
            { tier: 'Strong AoE — Hits Twice',     name: 'Thunderstorm',    pow: 85,  cost: 22, hits: 2, note: 'A relentless storm of lightning crashes into all enemies twice.' },
        ]
    },
    {
        element: 'Wind',
        spells: [
            { tier: 'Basic — Single Target',       name: 'Gust',            pow: 65,  cost: 6,  hits: 1, note: 'A sharp burst of wind against a single target.' },
            { tier: 'Powerful — Single Target',    name: 'Wind Cutter',     pow: 90,  cost: 13, hits: 1, note: 'A razor-edged wind blade that has a high natural accuracy bonus (+15%).' },
            { tier: 'AoE',                         name: 'Cyclone',         pow: 70,  cost: 16, hits: 1, note: 'A spiralling vortex sweeps through all enemies.' },
            { tier: 'Strong AoE — Hits Twice',     name: 'Tempest',         pow: 85,  cost: 22, hits: 2, note: 'A violent storm tears through all enemies twice. 10% chance per hit to reduce target EVA by 1 for 2 rounds.' },
        ]
    },
    {
        element: 'Earth',
        spells: [
            { tier: 'Basic — Single Target',       name: 'Stone Throw',     pow: 65,  cost: 6,  hits: 1, note: 'A jagged rock is hurled at a single target.' },
            { tier: 'Powerful — Single Target',    name: 'Boulder Crash',   pow: 90,  cost: 14, hits: 1, note: 'A massive rock is dropped onto a single target. 20% chance to reduce target DEF by 1 for 2 rounds.' },
            { tier: 'AoE',                         name: 'Earthquake',      pow: 70,  cost: 17, hits: 1, note: 'The ground trembles, striking all grounded enemies.' },
            { tier: 'Strong AoE — Hits Twice',     name: 'Landslide',       pow: 85,  cost: 24, hits: 2, note: 'A cascade of earth and rock buries all enemies twice.' },
        ]
    },
    {
        element: 'Grass',
        spells: [
            { tier: 'Basic — Single Target',       name: 'Spore',           pow: 65,  cost: 6,  hits: 1, note: 'A cloud of toxic spores is launched at a single target.' },
            { tier: 'Powerful — Single Target',    name: 'Thorn Whip',      pow: 90,  cost: 13, hits: 1, note: 'Barbed vines lash a single target. Restores 15% of damage dealt as HP to the caster.' },
            { tier: 'AoE',                         name: 'Petal Storm',     pow: 70,  cost: 16, hits: 1, note: 'A storm of razor-edged petals cuts through all enemies.' },
            { tier: 'Strong AoE — Hits Twice',     name: 'Overgrowth',      pow: 85,  cost: 22, hits: 2, note: 'An explosion of wild flora engulfs all enemies twice. Restores a small amount of HP to all allies after resolving.' },
        ]
    },
    {
        element: 'Dark',
        spells: [
            { tier: 'Basic — Single Target',       name: 'Shadow Bolt',     pow: 65,  cost: 6,  hits: 1, note: 'A bolt of dark energy against a single target.' },
            { tier: 'Powerful — Single Target',    name: 'Void Strike',     pow: 92,  cost: 14, hits: 1, note: 'A surge of void energy that reduces the target\'s ATK and MAG by 1 for 2 rounds on hit.' },
            { tier: 'AoE',                         name: 'Dark Pulse',      pow: 70,  cost: 16, hits: 1, note: 'A wave of dark energy radiates outward, hitting all enemies.' },
            { tier: 'Strong AoE — Hits Twice',     name: 'Oblivion',        pow: 85,  cost: 25, hits: 2, note: 'Crushing darkness descends on all enemies twice. Second hit always deals reduced damage (×0.75) to reflect diminishing returns.' },
        ]
    },
    {
        element: 'Light',
        spells: [
            { tier: 'Basic — Single Target',       name: 'Aura',            pow: 65,  cost: 6,  hits: 1, note: 'A basic Light spell against a single target. Already in game.' },
            { tier: 'Powerful — Single Target',    name: 'Holy Lance',      pow: 92,  cost: 14, hits: 1, note: 'A spear of holy light pierces a single target. Deals bonus damage against Dark-affinity enemies.' },
            { tier: 'AoE',                         name: 'Radiance',        pow: 70,  cost: 16, hits: 1, note: 'A blinding burst of holy light washes over all enemies.' },
            { tier: 'Strong AoE — Hits Twice',     name: 'Divine Wrath',    pow: 85,  cost: 24, hits: 2, note: 'A cascade of sacred energy crashes down on all enemies twice.' },
        ]
    },
    {
        element: 'Anima',
        spells: [
            { tier: 'Basic — Single Target',       name: 'Anima Bolt',      pow: 65,  cost: 6,  hits: 1, note: 'A pulse of raw Anima energy against a single target. No enemies are weak to Anima.' },
            { tier: 'Powerful — Single Target',    name: 'Soul Rend',       pow: 90,  cost: 14, hits: 1, note: 'Tears at a target\'s spiritual essence. Reduces target\'s next action damage by 20% (1 round).' },
            { tier: 'AoE',                         name: 'Supernova',       pow: 70,  cost: 12, hits: 1, note: 'An Anima spell that hits all opposing targets. Already in game.' },
            { tier: 'Strong AoE — Hits Twice',     name: 'Cosmic Collapse', pow: 85,  cost: 24, hits: 2, note: 'The fabric of reality tears twice, striking all enemies. Cannot be evaded by any means.' },
        ]
    },
];

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function cell(text, opts = {}) {
    const { bold = false, fill = 'FFFFFF', color = '000000', width = 2340, italic = false, center = false } = opts;
    return new TableCell({
        borders,
        width: { size: width, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: cellMargins,
        children: [new Paragraph({
            alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
            children: [new TextRun({ text, bold, color, italics: italic, font: 'Arial', size: 20 })]
        })]
    });
}

function headerRow(cols, fill, textColor = 'FFFFFF') {
    return new TableRow({
        tableHeader: true,
        children: cols.map(([text, width]) => new TableCell({
            borders,
            width: { size: width, type: WidthType.DXA },
            shading: { fill, type: ShadingType.CLEAR },
            margins: cellMargins,
            children: [new Paragraph({
                children: [new TextRun({ text, bold: true, color: textColor, font: 'Arial', size: 20 })]
            })]
        }))
    });
}

function spacer() {
    return new Paragraph({ children: [new TextRun({ text: '', size: 18 })] });
}

function heading1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text, font: 'Arial', size: 32, bold: true })]
    });
}

function heading2(text, color = '000000') {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text, font: 'Arial', size: 26, bold: true, color })]
    });
}

function bodyText(text) {
    return new Paragraph({
        children: [new TextRun({ text, font: 'Arial', size: 20 })]
    });
}

// ─── BUILD DOCUMENT ───────────────────────────────────────────────────────────
const children = [];

// Title
children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Spell Reference — Full Tier List', font: 'Arial', size: 40, bold: true, color: '1A1A2E' })]
}));
children.push(spacer());

// Intro
children.push(bodyText('This document catalogues all spells across 9 elements in a 4-tier structure. Each element has one spell per tier. Spells marked "Already in game" exist in the current PoC build.'));
children.push(spacer());

// Tier legend table
children.push(bodyText('Tier overview:'));
children.push(spacer());
children.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2600, 1200, 1200, 4360],
    rows: [
        headerRow([['Tier', 2600], ['Base POW', 1200], ['SP Cost', 1200], ['Description', 4360]], '2C3E50'),
        new TableRow({ children: [
            cell('Basic — Single Target',    { fill: 'F2F3F4', width: 2600 }),
            cell('65',                        { fill: 'F2F3F4', width: 1200, center: true }),
            cell('6–8',                       { fill: 'F2F3F4', width: 1200, center: true }),
            cell('A straightforward elemental spell. Low cost, reliable damage, single enemy.', { fill: 'F2F3F4', width: 4360 }),
        ]}),
        new TableRow({ children: [
            cell('Powerful — Single Target',  { width: 2600 }),
            cell('90–95',                     { width: 1200, center: true }),
            cell('12–16',                     { width: 1200, center: true }),
            cell('A stronger focused spell. Often carries a secondary effect.', { width: 4360 }),
        ]}),
        new TableRow({ children: [
            cell('AoE',                       { fill: 'F2F3F4', width: 2600 }),
            cell('70',                        { fill: 'F2F3F4', width: 1200, center: true }),
            cell('14–18',                     { fill: 'F2F3F4', width: 1200, center: true }),
            cell('Hits all enemies once. Lower POW to offset the multi-target advantage.', { fill: 'F2F3F4', width: 4360 }),
        ]}),
        new TableRow({ children: [
            cell('Strong AoE — Hits Twice',   { width: 2600 }),
            cell('85 ×2',                     { width: 1200, center: true }),
            cell('20–26',                     { width: 1200, center: true }),
            cell('Hits all enemies twice. High SP investment. Usually carries a bonus effect.', { width: 4360 }),
        ]}),
    ]
}));
children.push(spacer());
children.push(spacer());

// Per-element sections
for (const { element, spells } of SPELLS) {
    const col = COLS[element];
    children.push(heading2(`${element} Spells`, col.header));
    children.push(spacer());

    children.push(new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2200, 1800, 700, 700, 600, 3360],
        rows: [
            headerRow(
                [['Tier', 2200], ['Name', 1800], ['POW', 700], ['SP', 700], ['Hits', 600], ['Notes', 3360]],
                col.header
            ),
            ...spells.map((sp, i) => new TableRow({
                children: [
                    cell(sp.tier,  { fill: i % 2 === 0 ? col.light : 'FFFFFF', width: 2200, italic: true }),
                    cell(sp.name,  { fill: i % 2 === 0 ? col.light : 'FFFFFF', width: 1800, bold: true }),
                    cell(String(sp.pow),  { fill: i % 2 === 0 ? col.light : 'FFFFFF', width: 700, center: true }),
                    cell(String(sp.cost), { fill: i % 2 === 0 ? col.light : 'FFFFFF', width: 700, center: true }),
                    cell(sp.hits === 2 ? '×2' : '1', { fill: i % 2 === 0 ? col.light : 'FFFFFF', width: 600, center: true }),
                    cell(sp.note,  { fill: i % 2 === 0 ? col.light : 'FFFFFF', width: 3360 }),
                ]
            }))
        ]
    }));
    children.push(spacer());
    children.push(spacer());
}

// ─── PACK & WRITE ─────────────────────────────────────────────────────────────
const doc = new Document({
    styles: {
        default: { document: { run: { font: 'Arial', size: 20 } } },
        paragraphStyles: [
            { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
              run: { size: 40, bold: true, font: 'Arial', color: '1A1A2E' },
              paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
            { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
              run: { size: 26, bold: true, font: 'Arial' },
              paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
        ]
    },
    sections: [{
        properties: {
            page: {
                size: { width: 12240, height: 15840 },
                margin: { top: 1440, right: 1080, bottom: 1440, left: 1080 }
            }
        },
        children
    }]
});

Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync('/mnt/user-data/outputs/Spell_Reference.docx', buf);
    console.log('Done');
});