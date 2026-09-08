# Extrator de Relatórios – Apoena

Extrai dados de PDFs (Notas Fiscais/NAI, Diárias de Hotel, Exames Ocupacionais e
Mapa de Refeições) e gera uma planilha Excel formatada.

Os PDFs processados costumam conter dados pessoais e fiscais — nome, CPF, CNPJ,
e-mail, telefone. Por isso a forma recomendada de uso é **na própria máquina**:
nada sai do computador, nem para a nuvem, nem para terceiros.

---

## 1. Programa para Windows (recomendado)

Não precisa instalar Python, nem Tesseract, nem internet para funcionar.

1. Baixe o `ExtratorApoena-windows.zip`:
   - da aba **Actions** do repositório → execução mais recente de *Executável Windows* → **Artifacts**; ou
   - da página de **Releases**, quando houver uma versão publicada.
2. Extraia o `.zip` em uma pasta (por exemplo, `C:\ExtratorApoena`).
3. Dê dois cliques em **`ExtratorApoena.exe`**.

Abre uma janela preta (o servidor) e, em seguida, o navegador com o app. Para
encerrar, feche a janela preta.

O Tesseract vai embutido no pacote, com os idiomas português, inglês e detecção de
orientação — **o OCR de PDFs digitalizados funciona sem nenhuma instalação extra**.

Detalhes práticos:

- O programa é distribuído como pasta compactada, não como arquivo único: o
  Streamlit carrega recursos do disco, e um `.exe` único teria de descompactar
  centenas de MB a cada abertura.
- Na primeira execução o Windows pode pedir permissão de rede — o servidor roda
  só em `localhost`, sem expor nada para fora.
- O SmartScreen pode avisar que o programa é de origem desconhecida (ele não é
  assinado digitalmente). "Mais informações" → "Executar assim mesmo".

Para gerar o executável você mesmo, em uma máquina Windows:

```powershell
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean desktop/extrator.spec
# resultado em dist\ExtratorApoena\
```

## 2. Rodar a partir do código

```bash
pip install -r requirements.txt
streamlit run app_web.py          # ou: python desktop/launcher.py
```

Todas as dependências de `requirements.txt` são wheels Python — nenhum pacote de
sistema é obrigatório.

## 3. Servidor próprio (Docker)

Serve para Render, Fly.io, Railway, Hugging Face Spaces ou um servidor seu:

```bash
docker build -t extrator-apoena .
docker run -p 8501:8501 extrator-apoena
```

A imagem instala o Tesseract normalmente, então o OCR funciona. Ela respeita a
variável `PORT`, que essas plataformas definem sozinhas.

Lembre-se de que hospedar significa enviar os PDFs para fora da sua máquina: se
optar por este caminho, use um serviço privado, com acesso restrito.

---

## OCR

A leitura padrão usa a **camada de texto** do PDF — mais rápida e mais precisa
que o OCR. O OCR entra em PDFs digitalizados (imagem) ou como fallback, e depende
do binário do **Tesseract**:

```bash
sudo xargs -a apt-packages.txt apt-get install -y   # Linux: tesseract-ocr, tesseract-ocr-por
brew install tesseract tesseract-lang               # macOS
```

No Windows, o executável da seção 1 já traz tudo embutido.

Sem o Tesseract o app **continua funcionando**: a opção "OCR Alta Precisão"
aparece desabilitada e os arquivos que exigiriam OCR mostram um aviso explicando
o motivo, em vez de derrubar a execução.

## Por que não usamos a Streamlit Community Cloud

O deploy por lá quebrou por um defeito da imagem base deles: sobrou uma entrada de
`bullseye-security` no `sources.list` de uma imagem que já é trixie, e o `Release`
desse repositório expirou. Como a plataforma roda `apt-get update` sempre que
encontra um `packages.txt`, o passo de dependências passou a falhar:

```
E: Release file for http://deb.debian.org/debian-security/dists/bullseye-security/InRelease is expired
❗️ installer returned a non-zero exit code
```

Não havia correção possível pelo repositório — o `sources.list` é da
infraestrutura deles. O app segue preparado para rodar lá (não depende de nenhum
pacote apt), mas as opções acima são melhores: o executável não depende de
plataforma nenhuma, e o Docker nos dá controle da imagem base.

A lista de pacotes apt continua versionada em `apt-packages.txt`.
