<div>
  <h2>
    Projeto para automatizar processo - comparação de dados de produção e desenv.
  </h2>
  <h2>
    MÉTODOS
  </h2>
</div>
<div>
  <b>geraPastasTemplates</b> - Percorre o projeto FMC e buscas todos os formulários existentes e cria um arquivo de txt.<br><br>
    
  <b>geraUrlsProd</b> - monta a url de produção para baixar o .jar dos servidores e gera um arquivo .txt com as urls montadas.<br><br>

  <b>abrirURLNavegador</b> - abre o navegar e passa a url de produção e iniciar o download.<br><br>

  <b>copiaArquivoJar</b> - Percorre os arquivos .jar baixado, copia para uma nova pasta dentro do projeto chamada <b>formulários</b> (caso não exista, será criado) e também cria a subpasta para cada formulário. Caso não encontre o .jar, cria uma pasta urlErroNoDownload e passa os formulários que não foi possível baixar.<br><br>

  <b>extrairJarFormularioSubpasta</b> - Percorre os arquivos .jar copiado na pasta <b>formulários</b> e extrai o arquivo.<br><br>

  <b>copiarArquivosPastasEspecificas</b> - Verifica o tipo de arquivo e separa em pastas especificas criando duas novas pastas <b>PastaJasper</b> e <b>PastaQFS</b>.<br><br>

  <b>gerarPlanilhaExcel</b> - Pega os números de cada formulário em <b>PastaJasper</b> e <b>PastaQFS</b>, lista todos os formulários em suas expectativas colunas.<br><br>

  <b>verificarFormularios</b> - Compara e válida qual formulário está na tecnologia antiga (QFS) ou já foi migrado (Jasper) em relação a desenv e produção, passando a informação em uma coluna extra na planilha criada no método gerarPlanilhaExcel.<br><br>
</div>
