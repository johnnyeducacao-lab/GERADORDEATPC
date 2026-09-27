# Gerador de ATPC — EE Renata Graziano

Projeto completo para registrar ATPCs, gerar atas com IA, gerar DOCX com cabeçalho oficial e converter o próprio DOCX para PDF.

## Já vem cadastrado
- Johnny — quinta, 16:40
- Nailson — quinta, 19:00–20:30
- Vânia — quarta, 13:00
- Rosa — quarta, 10:40
- Listas de professores fornecidas pelo usuário, com CPF/RG quando informado.

## Funções
- Carrega automaticamente professores pelo coordenador.
- Edita nome e CPF/RG diretamente na ATPC.
- Adiciona ou remove professor da ATPC.
- Marca Presente, Ausente ou Justificado.
- Seleciona outros coordenadores participantes.
- Seleciona quem escreveu a ata.
- Campos separados para pauta, informações obrigatórias e encaminhamentos.
- Geração de ata via OpenAI quando `OPENAI_API_KEY` estiver configurada.
- Fallback local se a chave não estiver configurada.
- Geração de DOCX com o cabeçalho fornecido.
- Conversão DOCX → PDF com LibreOffice.
- Histórico salvo no navegador (localStorage).

## Rodar localmente
Requer Node.js 20+. Para PDF, instale LibreOffice.

```bash
npm install
npm start
```
Abra `http://localhost:3000`.

## Publicar no Render
O projeto já inclui `Dockerfile` e `render.yaml`.

1. Crie um repositório no GitHub e envie estes arquivos.
2. No Render, escolha **New > Blueprint** ou **Web Service** apontando para o repositório.
3. Use Docker.
4. Adicione a variável `OPENAI_API_KEY` nas Environment Variables.
5. Faça o deploy.

> Render é recomendado para esta versão porque o PDF é realmente convertido a partir do DOCX usando LibreOffice. Em hospedagens serverless comuns, como Vercel, o LibreOffice não fica disponível da mesma forma.

## Segurança
O projeto atual é um protótipo administrativo. Antes de disponibilizar publicamente dados com CPF/RG, recomenda-se adicionar login e controle de acesso.
