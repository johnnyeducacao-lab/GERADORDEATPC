const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');
const util = require('util');
const execFileAsync = util.promisify(execFile);
const OpenAI = require('openai');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, Header, ImageRun, HeadingLevel,
  VerticalAlign
} = require('docx');

const app = express();
app.use(express.json({ limit: '3mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const headerPath = path.join(__dirname, 'public', 'assets', 'cabecalho-atpc.png');

function clean(v='') { return String(v ?? '').trim(); }
function safeFile(v='ATPC') { return clean(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-zA-Z0-9_-]+/g,'_'); }

function fallbackAta(data) {
  const presentes = (data.professores || []).filter(p => p.situacao === 'Presente').map(p => p.nome);
  const ausentes = (data.professores || []).filter(p => p.situacao === 'Ausente').map(p => p.nome);
  const justificados = (data.professores || []).filter(p => p.situacao === 'Justificado').map(p => p.nome);
  const coords = [data.coordenadorPrincipal, ...(data.coordenadoresParticipantes || [])].filter(Boolean);
  const partes = [];
  partes.push(`Aos ${clean(data.dataExtenso) || clean(data.data)}, realizou-se a Aula de Trabalho Pedagógico Coletivo (ATPC) da E.E. Renata Graziano de Oliveira Prado, no horário de ${clean(data.horarioInicio)}${data.horarioFim ? ` às ${clean(data.horarioFim)}` : ''}, sob a condução de ${coords.join(' e ') || 'coordenação pedagógica'}.`);
  if (clean(data.pauta)) partes.push(`A formação teve como pauta: ${clean(data.pauta).replace(/\n+/g, '; ')}.`);
  if (clean(data.informacoes)) partes.push(`Durante o encontro, foram discutidos os seguintes pontos: ${clean(data.informacoes).replace(/\n+/g, '; ')}.`);
  if (clean(data.encaminhamentos)) partes.push(`Como encaminhamentos e combinados, registrou-se: ${clean(data.encaminhamentos).replace(/\n+/g, '; ')}.`);
  if (presentes.length) partes.push(`Participaram do encontro ${presentes.length} professor(es), conforme lista de presença e assinaturas constantes neste registro.`);
  if (ausentes.length || justificados.length) partes.push(`Foram registradas ${ausentes.length} ausência(s) e ${justificados.length} ausência(s) justificada(s).`);
  if (clean(data.redator)) partes.push(`O registro desta ata ficou sob responsabilidade de ${clean(data.redator)}.`);
  partes.push('Nada mais havendo a registrar, o encontro foi encerrado, ficando os participantes cientes das orientações e encaminhamentos apresentados.');
  return partes.join('\n\n');
}

app.post('/api/gerar-ata', async (req, res) => {
  const data = req.body || {};
  const fallback = fallbackAta(data);
  if (!process.env.OPENAI_API_KEY) return res.json({ ata: fallback, modo: 'modelo-local' });
  try {
    const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
    const prompt = `Você redige atas de ATPC de uma escola pública estadual de São Paulo.\n\nREGRAS:\n- Produza uma ata formal, clara, objetiva e pedagógica.\n- NÃO invente acontecimentos, decisões, nomes, leis, falas ou informações.\n- Use obrigatoriamente todos os pontos de \"Informações que não podem faltar\".\n- Diferencie pauta planejada do que efetivamente foi discutido.\n- Inclua encaminhamentos quando fornecidos.\n- Cite coordenador principal, coordenadores participantes, data, horário e responsável pela redação.\n- Não liste CPFs/RGs no corpo da ata.\n- Não mencione que foi gerada por IA.\n\nDADOS:\n${JSON.stringify(data, null, 2)}\n\nEscreva somente o texto final da ata.`;
    const response = await client.responses.create({
      model: process.env.OPENAI_MODEL || 'gpt-5.6',
      input: prompt
    });
    const text = response.output_text?.trim();
    res.json({ ata: text || fallback, modo: text ? 'ia' : 'modelo-local' });
  } catch (e) {
    console.error(e);
    res.json({ ata: fallback, modo: 'modelo-local', aviso: 'IA indisponível; foi usada a redação local.' });
  }
});

function noBorders() {
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };
}

function cell(text, opts={}) {
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text: clean(text), bold: !!opts.bold, size: opts.size || 20 })]
    })]
  });
}

function buildDocx(data) {
  const headerImage = fs.readFileSync(headerPath);
  const coordParticipantes = (data.coordenadoresParticipantes || []).filter(Boolean);
  const professores = data.professores || [];
  const rows = [
    new TableRow({ children: [
      cell('Nº', { bold: true, width: 6, align: AlignmentType.CENTER }),
      cell('NOME DO PROFESSOR', { bold: true, width: 42, align: AlignmentType.CENTER }),
      cell('CPF OU RG', { bold: true, width: 18, align: AlignmentType.CENTER }),
      cell('SITUAÇÃO', { bold: true, width: 14, align: AlignmentType.CENTER }),
      cell('ASSINATURA', { bold: true, width: 20, align: AlignmentType.CENTER })
    ]})
  ];
  professores.forEach((p, i) => rows.push(new TableRow({ children: [
    cell(String(i+1), { align: AlignmentType.CENTER }),
    cell(p.nome),
    cell(p.documento || ''),
    cell(p.situacao || 'Presente', { align: AlignmentType.CENTER }),
    cell('________________________', { align: AlignmentType.CENTER })
  ]})));

  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'REGISTRO DE ATPC', bold: true, size: 28 })] }),
    new Paragraph({ children: [] }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
      new TableRow({ children: [cell(`Data: ${clean(data.data)}`, { bold: true, width: 35 }), cell(`Horário: ${clean(data.horarioInicio)}${data.horarioFim ? ' às ' + clean(data.horarioFim) : ''}`, { bold: true, width: 35 }), cell(`Dia: ${clean(data.diaSemana)}`, { bold: true, width: 30 })] }),
      new TableRow({ children: [cell(`Coordenador(a) responsável: ${clean(data.coordenadorPrincipal)}`, { bold: true, width: 50 }), cell(`Responsável pela ata: ${clean(data.redator)}`, { bold: true, width: 50 })] }),
      new TableRow({ children: [cell(`Outros coordenadores participantes: ${coordParticipantes.length ? coordParticipantes.join(', ') : '—'}`, { width: 100 })] })
    ]}),
    new Paragraph({ children: [] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'PAUTA', bold: true })] }),
    new Paragraph({ children: [new TextRun({ text: clean(data.pauta) || '—', size: 22 })] }),
    new Paragraph({ children: [] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'ATA DO ATPC', bold: true })] }),
    ...clean(data.ata).split(/\n\n+/).map(t => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 150 }, children: [new TextRun({ text: t, size: 22 })] })),
  ];
  if (clean(data.encaminhamentos)) {
    children.push(new Paragraph({ children: [] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'ENCAMINHAMENTOS / COMBINADOS', bold: true })] }));
    children.push(new Paragraph({ children: [new TextRun({ text: clean(data.encaminhamentos), size: 22 })] }));
  }
  children.push(new Paragraph({ children: [] }));
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'LISTA DE PRESENÇA', bold: true })] }));
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows }));
  children.push(new Paragraph({ children: [] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 450 }, children: [new TextRun({ text: '____________________________________________', size: 22 })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: clean(data.coordenadorPrincipal) || 'Coordenação Pedagógica', bold: true, size: 20 })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Coordenação Pedagógica', size: 18 })] }));

  return new Document({
    sections: [{
      properties: { page: { margin: { top: 2200, right: 900, bottom: 900, left: 900 } } },
      headers: {
        default: new Header({
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: headerImage, transformation: { width: 640, height: 209 }, type: 'png' })] })]
        })
      },
      children
    }]
  });
}

app.post('/api/docx', async (req, res) => {
  try {
    const data = req.body || {};
    const doc = buildDocx(data);
    const buffer = await Packer.toBuffer(doc);
    const name = `ATPC_${safeFile(data.coordenadorPrincipal)}_${safeFile(data.data)}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${name}"`);
    res.send(buffer);
  } catch (e) {
    console.error(e); res.status(500).json({ error: 'Falha ao gerar DOCX.' });
  }
});

app.post('/api/pdf', async (req, res) => {
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'atpc-'));
  try {
    const data = req.body || {};
    const doc = buildDocx(data);
    const buffer = await Packer.toBuffer(doc);
    const base = `ATPC_${safeFile(data.coordenadorPrincipal)}_${safeFile(data.data)}`;
    const docxPath = path.join(work, `${base}.docx`);
    fs.writeFileSync(docxPath, buffer);
    await execFileAsync('libreoffice', ['--headless', '--convert-to', 'pdf', '--outdir', work, docxPath], { timeout: 60000 });
    const pdfPath = path.join(work, `${base}.pdf`);
    if (!fs.existsSync(pdfPath)) throw new Error('PDF não criado');
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${base}.pdf"`);
    res.send(fs.readFileSync(pdfPath));
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Falha ao converter DOCX para PDF. No deploy, use o Dockerfile incluído (LibreOffice).' });
  } finally {
    setTimeout(() => fs.rm(work, { recursive: true, force: true }, () => {}), 2000);
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`ATPC rodando em http://localhost:${port}`));
