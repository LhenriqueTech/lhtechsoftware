# Gerador de Relatórios - LH TECH

## Requisitos
- Windows (recomendado) ou Linux/Mac com Python 3.9+
- Python e pip instalados

## Instalação (modo desenvolvimento)
1. Copie todos os arquivos para uma pasta, por exemplo `lhtech_relatorios`.
2. Abra terminal nesta pasta.
3. Rode: pip install -r requirements.txt
python app.py

4. A aplicação abrirá. Selecione o arquivo base (o `.xlsx` que você usa) e clique em Gerar.

## Gerar .exe (Windows)
1. Abra `build_exe.bat` (ou rode manualmente):
pyinstaller --onefile --windowed --add-data "logo.png;." --name "GeradorRelatoriosLHTech" app.py

2. Executável será gerado em `dist\GeradorRelatoriosLHTech.exe`.

## Observações
- O app **não altera** o arquivo base. Ele gera um arquivo `relatorios_modelo_todos_gerado.xlsx` na pasta de saída escolhida.
- A planilha final (template) foi respeitada — fórmulas e estilos essenciais são aplicados conforme seu modelo final.
- Se sua localidade Excel exige `SE()` em vez de `IF()`, as fórmulas de extras podem mostrar `#NOME?` no Excel local. Se ocorrer, altere a linha em `processor.py`:
```py
ws.cell(row=row, column=10, value=f"=SE(H{row}>I{row}; H{row}-I{row}; 0)")
(trocar a função IF por SE e os separadores , por ; conforme sua localidade.)


---

## Notas finais / recomendações rápidas

1. **Testes locais**: execute `python app.py`, carregue o arquivo base (`Relatório de atendimento_...xlsx`) e gere o relatório. Confira a aba **ALINE** gerada. Se o Excel no seu PC mostrar `#NOME?` em `IF`, troque a fórmula para `SE` (veja README).
2. **Ícone do `.exe`**: se quiser, forneça um `ico` e eu atualizo o `pyinstaller` command (`--icon=icone.ico`).
3. **Personalizações UX**: posso adicionar opção de tema escuro, lembrar diretório de saída (arquivo JSON de preferências) e mostrar pré-visualização em tabela dentro do app. Diga qual recurso quer a seguir.
4. **Garantia**: o `processor.py` tem a mesma lógica que estávamos usando e respeita a planilha final — não faz mudanças inesperadas.

---

Se quiser, já gero:
- (A) uma versão rápida do `.exe` aqui (não posso criar arquivos .exe fora do ambiente do usuário) — eu te dou o comando exato (`pyinstaller`) e config; ou
- (B) eu já converto todas as fórmulas `IF` para `SE` automaticamente dependendo da localidade detectada (posso adicionar detecção automática do Excel locale no código).

Quer que eu:
1. já adapte automaticamente as fórmulas `IF -> SE` no `processor.py` conforme a localidade do Windows (adiciono essa lógica), **ou**
2. deixe como está e você fará a troca manual se surgir o `#NOME?`?

Também posso já incluir o ícone `.ico` no build se você enviar o arquivo (ou eu uso o `logo.png` convertido).

