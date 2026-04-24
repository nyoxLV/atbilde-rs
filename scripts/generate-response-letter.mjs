import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  LevelFormat,
  convertInchesToTwip,
} from "docx"
import fs from "node:fs"
import path from "node:path"

const FONT = "Times New Roman"

function p(text, opts = {}) {
  const {
    bold = false,
    italic = false,
    size = 22, // 11pt
    alignment = AlignmentType.JUSTIFIED,
    spacingAfter = 120,
    spacingBefore = 0,
    indent,
  } = opts
  return new Paragraph({
    alignment,
    spacing: { after: spacingAfter, before: spacingBefore, line: 300 },
    indent,
    children: [
      new TextRun({
        text,
        bold,
        italics: italic,
        font: FONT,
        size,
      }),
    ],
  })
}

function mixed(runs, opts = {}) {
  const { alignment = AlignmentType.JUSTIFIED, spacingAfter = 120, spacingBefore = 0, indent } = opts
  return new Paragraph({
    alignment,
    spacing: { after: spacingAfter, before: spacingBefore, line: 300 },
    indent,
    children: runs.map(
      (r) =>
        new TextRun({
          text: r.text,
          bold: !!r.bold,
          italics: !!r.italic,
          font: FONT,
          size: r.size ?? 22,
        }),
    ),
  })
}

function heading(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120, line: 300 },
    children: [new TextRun({ text, bold: true, font: FONT, size: 24 })],
  })
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80, line: 300 },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, font: FONT, size: 22 })],
  })
}

function numbered(text) {
  return new Paragraph({
    numbering: { reference: "numbered", level: 0 },
    spacing: { after: 80, line: 300 },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, font: FONT, size: 22 })],
  })
}

const children = []

// Header block – sender
children.push(
  p("SIA “UEC Sistēmas”", { bold: true, alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("Vienotais reģistrācijas Nr. 42103092124", { alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("E-pasts: janis.liepins@uecsistemas.lv", { alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("Tālr.: +371 26 521 145", { alignment: AlignmentType.LEFT, spacingAfter: 240 }),
)

// Recipient
children.push(
  p("RP SIA “Rīgas satiksme”", { bold: true, alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("Valdes priekšsēdētājai Dž. Innusas kundzei", { alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("Kleistu iela 28, Rīga, LV-1067", { alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("Kopija: agnese.mengele@rigassatiksme.lv", { alignment: AlignmentType.LEFT, spacingAfter: 360 }),
)

// Date / place
children.push(
  p("Rīgā, 2026. gada ___. aprīlī", { alignment: AlignmentType.RIGHT, spacingAfter: 240 }),
)

// Subject
children.push(
  mixed(
    [
      { text: "Par: ", bold: true },
      {
        text:
          "Atbildi uz 2026. gada aprīļa vēstuli par līgumu Nr. LIG-IEP/2025/172 un Nr. LIG-IEP/2025/173 izpildi un aprēķinātajiem zaudējumiem EUR 2520,15 apmērā",
        bold: true,
      },
    ],
    { alignment: AlignmentType.LEFT, spacingAfter: 360 },
  ),
)

children.push(p("Godājamā Innusas kundze,", { alignment: AlignmentType.LEFT, spacingAfter: 180 }))

children.push(
  mixed(
    [
      { text: "SIA “UEC Sistēmas” (turpmāk – " },
      { text: "Izpildītājs", bold: true },
      {
        text:
          ") ir saņēmusi RP SIA “Rīgas satiksme” (turpmāk – ",
      },
      { text: "Pasūtītājs", bold: true },
      {
        text:
          ") 2026. gada aprīļa vēstuli, kurā Pasūtītājs informē par aprēķinātajiem zaudējumiem EUR 2520,15 apmērā saistībā ar līgumu Nr. LIG-IEP/2025/172 un Nr. LIG-IEP/2025/173 (turpmāk – ",
      },
      { text: "Līgumi", bold: true },
      { text: ") izpildi." },
    ],
    { spacingAfter: 180 },
  ),
)

children.push(
  mixed(
    [
      { text: "Izpildītājs ar cieņu " },
      { text: "nepiekrīt", bold: true },
      {
        text:
          " Pasūtītāja aprēķinātajai zaudējumu summai un tās pamatojumam turpmāk minēto iemeslu dēļ.",
      },
    ],
    { spacingAfter: 120 },
  ),
)

// 1
children.push(heading("1. Par faktiskajiem apstākļiem – invertori ir ražojuši elektroenerģiju"))
children.push(
  mixed(
    [
      { text: "Pretēji vēstulē norādītajam, abās adresēs (" },
      { text: "Vestienas ielā 35", bold: true },
      { text: " un " },
      { text: "Brīvības ielā 191, Rīgā", bold: true },
      {
        text:
          ") uzstādītās saules elektrostacijas invertori darbojas un ir ražojuši elektroenerģiju kopš 2026. gada janvāra. Tas attiecīgi samazina Pasūtītāja no ārējiem piegādātājiem iepirkto elektroenerģijas apjomu un izmaksas jau kopš minētā brīža.",
      },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  mixed(
    [
      {
        text:
          "Apstāklis, ka objekts formāli nav nodots ar abpusēji parakstītu Darbu pieņemšanas un nodošanas aktu (Līgumu 4. pielikums), neatspēko faktu, ka ",
      },
      {
        text: "Pasūtītājs jau faktiski saņem Līguma priekšmeta sniegto ekonomisko labumu",
        bold: true,
      },
      { text: " – elektroenerģiju pašpatēriņam." },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  mixed(
    [
      {
        text:
          "Izpildītāja puses pārstāvji atkārtoti informēja Pasūtītāja darbiniekus par to, ka invertori darbojas un ražo elektroenerģiju; atlikušās atklātās nepilnības (tostarp attālinātās monitoringa datu pārraides izveide un atsevišķi blakus konstatētie trūkumi) ",
      },
      {
        text: "neietekmēja invertoru darbības spēju un faktisko elektroenerģijas ražošanu",
        bold: true,
      },
      { text: "." },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  mixed(
    [
      {
        text:
          "Līdz ar to Pasūtītāja aprēķinā par periodu 2026. gada janvāris, februāris un marts norādītā “neražotā elektroenerģija” (177 + 482 + 12 471 = 13 130 kWh jeb EUR 1677,80 no kopējās prasītās summas) ",
      },
      { text: "neatbilst faktiskajai situācijai", bold: true },
      {
        text:
          " un nevar tikt uzskatīta par Izpildītāja rīcības rezultātā radītiem zaudējumiem.",
      },
    ],
    { spacingAfter: 120 },
  ),
)

// 2
children.push(heading("2. Par zaudējumu aprēķina metodoloģiju"))
children.push(
  p(
    "Pasūtītājs vēstulē pats norāda, ka zaudējumi ir aprēķināti, balstoties uz citu, jau uzstādītu līdzvērtīgas jaudas elektrostaciju saražotās elektroenerģijas apmēru, nevis uz faktiskiem mērījumiem konkrētajos objektos.",
  ),
)
children.push(
  mixed(
    [
      { text: "Saskaņā ar " },
      { text: "Civillikuma 1770. pantu", bold: true },
      { text: " zaudējums ir katrs mantiski novērtējams pametums, un saskaņā ar " },
      { text: "Civillikuma 1779. pantu", bold: true },
      {
        text:
          " zaudējumi ir atlīdzināmi tikai tad, ja tie ir tiešā cēloniskā sakarā ar vainojamu rīcību vai bezdarbību un ir pienācīgi pierādīti. Pierādīšanas pienākums gulstas uz to Pusi, kas prasa zaudējumu atlīdzību (",
      },
      { text: "Civilprocesa likuma 93. pants", bold: true },
      { text: ")." },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  p(
    "Aprēķins, kas balstīts uz pieņēmumu par “iespējamo ražošanu” pie citiem objektiem, nav uzskatāms par pierādītiem faktiskiem zaudējumiem – tas ir hipotētisks atrautās peļņas aprēķins, kas prasa ievērojami augstāku pierādīšanas standartu. Pasūtītājs nav iesniedzis:",
  ),
)
children.push(
  bullet(
    "objektu tehnisko salīdzināmību apliecinošus datus (jumta orientācija, slīpums, ēnojums, kopējā uzstādītā jauda, invertoru specifikācija);",
  ),
  bullet("faktisko meteoroloģisko apstākļu korekciju attiecīgajā periodā;"),
  bullet(
    "aprēķina pamatā esošos dokumentus, kurus Izpildītājs varētu pārbaudīt.",
  ),
)
children.push(p("Šādos apstākļos aprēķinu nav iespējams uzskatīt par tiesiski pamatotu."))

// 3
children.push(heading("3. Par kavējuma iemesliem un Pušu sadarbības pienākumu"))
children.push(
  mixed(
    [
      { text: "Atgādinām, ka saskaņā ar Līgumu " },
      { text: "8. nodaļu (Kavējumi un termiņu pagarinājumi)", bold: true },
      {
        text:
          " Pusēm ir pienākums savstarpēji saskaņot rīcību kavējumu gadījumā, un Pasūtītājam ir pienākums izskatīt Izpildītāja iesniegtos paskaidrojumus noteiktā termiņā.",
      },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  mixed(
    [
      {
        text:
          "Izpildītājs ir gatavs iesniegt detalizētu pārskatu par faktiskajiem kavējuma iemesliem, tostarp par apstākļiem, kas ",
      },
      { text: "nebija atkarīgi vienīgi no Izpildītāja", bold: true },
      { text: ", piemēram:" },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  bullet("piekļuves un saskaņošanas jautājumi Objektos;"),
  bullet(
    "tehnisko pieslēgumu un trešo personu (t.sk. sadales tīkla operatora) procedūru ilgums;",
  ),
  bullet("blakus konstatētie trūkumi, par kuriem Puses komunicēja darba kārtībā."),
)
children.push(
  mixed(
    [
      { text: "Saskaņā ar " },
      { text: "Civillikuma 1587. un 1588. pantu", bold: true },
      {
        text:
          " līgumi izpildāmi labā ticībā, un Pasūtītāja pienākums ir arī mazināt iespējamos zaudējumus (Civillikuma 1776. pants un tā piemērošanas prakse). Ja invertori faktiski darbojās, bet objekts netika formāli pieņemts, Pasūtītāja ekonomiskais zaudējums šajā daļā neiestājās.",
      },
    ],
    { spacingAfter: 120 },
  ),
)

// 4
children.push(heading("4. Par Līgumu 12.2. punkta piemērošanu"))
children.push(
  mixed(
    [
      { text: "Līgumu 12.2. punkts paredz atbildību par " },
      { text: "“nodarītiem”", bold: true },
      {
        text:
          " tiešiem un netiešiem zaudējumiem. Šī norma nerada Pasūtītājam tiesības vienpusēji noteikt zaudējumu apmēru bez pierādīšanas un cēloņsakarības konstatēšanas. Tā vienīgi apstiprina vispārējos Civillikuma principus par zaudējumu atlīdzināšanu, kas prasa:",
      },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  numbered("pretlikumīgu (līgumam neatbilstošu) rīcību vai bezdarbību;"),
  numbered("faktiskus un pierādītus zaudējumus;"),
  numbered("cēloņsakarību starp rīcību un zaudējumiem;"),
  numbered("Izpildītāja vainu."),
)
children.push(
  p(
    "Vismaz 2. un 3. elements Pasūtītāja aprēķinā nav pierādīts.",
  ),
)

// 5
children.push(heading("5. Par līgumsodu par termiņu neievērošanu"))
children.push(
  mixed(
    [
      { text: "Pievēršam uzmanību, ka Līgumi " },
      {
        text: "jau paredz speciālu tiesisko seku mehānismu par termiņu neievērošanu",
        bold: true,
      },
      {
        text:
          " – līgumsodu 0,1 % apmērā no Līguma cenas par katru nokavēto dienu, bet ne vairāk kā 10 % no Līguma cenas (Līgumu 11. nodaļa). Tas ir Pušu iepriekš nolīgtais un samērīgais tiesiskais risinājums kavējuma gadījumā. Papildu “atrautās peļņas” tipa zaudējumu prasīšana, kas pārklājas ar līgumsoda funkciju, rada dubultas atbildības risku un pieprasa īpaši stingru pierādīšanas standartu.",
      },
    ],
    { spacingAfter: 120 },
  ),
)

// 6
children.push(heading("6. Izpildītāja priekšlikumi turpmākai rīcībai"))
children.push(
  mixed(
    [
      { text: "Ņemot vērā augstāk minēto, Izpildītājs " },
      { text: "noraida", bold: true },
      {
        text:
          " vēstulē aprēķināto zaudējumu summu EUR 2520,15 apmērā kā nepamatotu un neatbilstošu faktiskajai situācijai.",
      },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  mixed(
    [
      { text: "Vienlaikus Izpildītājs apstiprina " },
      { text: "gatavību konstruktīvi sadarboties", bold: true },
      { text: ", lai pēc iespējas īsākā termiņā:" },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  numbered(
    "pabeigtu visus atlikušos darbus un noformētu abpusēji parakstītu Darbu pieņemšanas un nodošanas aktu abās adresēs;",
  ),
  numbered(
    "organizētu kopīgu objekta apskati un invertoru darbības pārbaudi (t.sk. invertoru iekšējo darbības žurnālu (event log) un saražotās enerģijas datu fiksāciju), lai objektīvi konstatētu faktisko ražošanas periodu un apjomu;",
  ),
  numbered(
    "iesniegtu rakstisku skaidrojumu par kavējuma iemesliem saskaņā ar Līgumu 8. nodaļu;",
  ),
  numbered(
    "vienotos par EKII finansējuma nosacījumu izpildei nepieciešamo darbību grafiku, lai izslēgtu riskus Pasūtītājam.",
  ),
)
children.push(
  mixed(
    [
      { text: "Lūdzam Pasūtītāju " },
      {
        text: "14 (četrpadsmit) dienu laikā",
        bold: true,
      },
      { text: " no šīs vēstules saņemšanas:" },
    ],
    { spacingAfter: 120 },
  ),
)
children.push(
  bullet(
    "iesniegt zaudējumu aprēķina detalizētu pamatojumu un pierādījumus (tostarp salīdzinājuma objektu datus, faktiskos mērījumus un aprēķina metodiku);",
  ),
  bullet("apstiprināt kopīgās objekta apskates datumu un laiku."),
)
children.push(
  p(
    "Izpildītājs patur visas Līgumos un normatīvajos aktos noteiktās tiesības, tostarp tiesības aizstāvēties pret jebkādām nepamatotām prasībām Latvijas Republikas tiesā.",
    { spacingAfter: 360 },
  ),
)

// Signature block
children.push(p("Ar cieņu,", { alignment: AlignmentType.LEFT, spacingAfter: 480 }))
children.push(
  p("Jānis Liepiņš", { bold: true, alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("Valdes loceklis", { alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("SIA “UEC Sistēmas”", { alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("E-pasts: janis.liepins@uecsistemas.lv", { alignment: AlignmentType.LEFT, spacingAfter: 0 }),
  p("Tālr.: +371 26 521 145", { alignment: AlignmentType.LEFT, spacingAfter: 240 }),
)
children.push(
  p("Dokuments parakstīts ar drošu elektronisko parakstu un satur laika zīmogu.", {
    italic: true,
    alignment: AlignmentType.LEFT,
  }),
)

const doc = new Document({
  creator: "SIA UEC Sistēmas",
  title: "Atbilde RP SIA Rīgas satiksme",
  description: "Atbildes vēstule par līgumiem LIG-IEP/2025/172 un LIG-IEP/2025/173",
  styles: {
    default: {
      document: {
        run: { font: FONT, size: 22 },
      },
    },
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: {
                indent: {
                  left: convertInchesToTwip(0.35),
                  hanging: convertInchesToTwip(0.25),
                },
              },
            },
          },
        ],
      },
      {
        reference: "numbered",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: {
                indent: {
                  left: convertInchesToTwip(0.4),
                  hanging: convertInchesToTwip(0.25),
                },
              },
            },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          margin: {
            top: convertInchesToTwip(1),
            bottom: convertInchesToTwip(1),
            left: convertInchesToTwip(1),
            right: convertInchesToTwip(1),
          },
        },
      },
      children,
    },
  ],
})

const outDir = "/vercel/share/v0-project"
const outPath = path.join(outDir, "Atbilde_RigasSatiksmei_UEC_Sistemas.docx")
const buffer = await Packer.toBuffer(doc)
fs.writeFileSync(outPath, buffer)
console.log("[v0] Wrote", outPath, buffer.length, "bytes")
