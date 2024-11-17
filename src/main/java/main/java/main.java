package main.java.main.java;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class main {

	public static void main(String[] args) throws InterruptedException, IOException {

		String diretorioProjeto = System.getProperty("user.dir");

		String diretorioBase = "C:\\kdi\\workspaceGit\\fmc-templates"; // Defina o caminho para o diretório que contém
																		// as pastas.
		String arquivoPastas = diretorioProjeto + "/pastas.txt";
		String arquivoUrls = diretorioProjeto + "/urls_Producao.txt";
		String pastaFormulariosProd = diretorioProjeto + "/formularios";
		String urlBase = "http://www12.bb.com.br/repov1/";
		String diretorioDownload = "C:\\Users\\C1332265\\Downloads";

		String pastaErros = diretorioProjeto + "/urlErroNoDownload";
			
			List<String> nomeDasPastas = Files.list(Paths.get(diretorioBase)).filter(Files::isDirectory) // Filtrar apenas
					// diretórios
					.map(Path::getFileName).map(Path::toString).collect(Collectors.toList());
			
			Files.createDirectories(Paths.get(pastaFormulariosProd));
			Files.createDirectories(Paths.get(pastaErros));
			
			//FUNÇÕES DO BOT
			
//			geraPastasTemplates(diretorioBase, arquivoPastas);
//			geraUrlsProd(arquivoPastas, arquivoUrls, urlBase, pastaFormulariosProd);
//			abrirURLNavegador(arquivoUrls);
//			copiaArquivoJar(nomeDasPastas, diretorioBase, diretorioDownload, pastaFormulariosProd, urlBase, pastaErros);
//			extrairJarFormularioSubpasta(nomeDasPastas, pastaFormulariosProd);
//			copiarArquivosPastasEspecificas(pastaFormulariosProd, diretorioProjeto);
//			gerarPlanilhaExcel(nomeDasPastas,pastaFormulariosProd, diretorioProjeto);
//			verificarFormularios(pastaFormulariosProd, diretorioBase, diretorioProjeto);

		System.out.println("!!!!Programa fechando, operação concluida!!!!");
	}
	
	// Listar todas as pastas no diretório especificado.
	private static void geraPastasTemplates(List<String> nomeDasPastas, String diretorioBase, String arquivoPastas) throws IOException {
		// Escrever o nome de cada pasta em um arquivo de texto.
		try (BufferedWriter writer = new BufferedWriter(new FileWriter(arquivoPastas))) {
			for (String pasta : nomeDasPastas) {
				writer.write(pasta);
				writer.newLine();
			}
		}
	}
	
	private static void geraUrlsProd(String arquivoPastas, String arquivoUrls, String urlBase, String pastaFormulariosProd) throws IOException {
		
		// Ler o arquivo de pastas e gerar URLs, escrevendo em um novo arquivo.
		try (BufferedReader reader = new BufferedReader(new FileReader(arquivoPastas));
				BufferedWriter writer = new BufferedWriter(new FileWriter(arquivoUrls))) {
			String linha;
			while ((linha = reader.readLine()) != null) {
				String urlCompleta = urlBase + linha + ".jar";
				writer.write(urlCompleta);
				writer.newLine();

				Files.createDirectories(Paths.get(pastaFormulariosProd, linha));
				System.out.println("Url criada: " + urlCompleta);
			}
		}
	}

	// Método para abrir uma URL no edge
	private static void abrirURLNavegador(String arquivoUrls) throws InterruptedException {
		// Ler o arquivo de URLs e abrir cada uma no Edge para download
		try (BufferedReader reader = new BufferedReader(new FileReader(arquivoUrls))) {
			while ((arquivoUrls = reader.readLine()) != null) {
				String comando = "cmd.exe /c start microsoft-edge:" + arquivoUrls;
				Runtime.getRuntime().exec(comando);
				System.out.println("Abrindo no Microsoft Edge: " + arquivoUrls);
				Thread.sleep(5000);
			}
		}
		catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	// Copiar os arquivos .jar para suas respectivas subpastas em nova_pasta
	private static void copiaArquivoJar(List<String> nomeDasPastas, String diretorioBase, String diretorioDownload, String pastaFormulariosProd,
			String urlBase, String pastaErros) throws IOException {
		
		for (String pasta : nomeDasPastas) {
			
			String nomeArquivoJar = pasta + ".jar";
			Path arquivoJarOrigem = Paths.get(diretorioDownload, nomeArquivoJar);
			Path subpastaDestino = Paths.get(pastaFormulariosProd, pasta);
			
		    try {
		        
		        if (Files.exists(arquivoJarOrigem)) {
		        	
		            // Verificar se a subpasta de destino existe, se não, criar
		            if (!Files.exists(subpastaDestino)) {
		                Files.createDirectories(subpastaDestino);
		            }
		            Path destino = subpastaDestino.resolve(nomeArquivoJar);
		            
		            Files.copy(arquivoJarOrigem, destino, StandardCopyOption.REPLACE_EXISTING);
		            System.out.println("Arquivo " + nomeArquivoJar + " copiado para " + destino.toString());
		        } else {
		            // Arquivo .jar não encontrado, criar arquivos de erro
		            System.out.println("Arquivo " + nomeArquivoJar + " não encontrado no diretório de downloads.");

		            // Verificar se a subpasta de destino existe, se não, criar
		            if (!Files.exists(subpastaDestino)) {
		                Files.createDirectories(subpastaDestino);
		            }

		            // Criar o arquivo de erro na subpasta específica
		            Path arquivoErroSubpasta = subpastaDestino.resolve("Não foi possível baixar.txt");
		            try (BufferedWriter writerErro = new BufferedWriter(new FileWriter(arquivoErroSubpasta.toFile()))) {
		                writerErro.write("Não foi possível baixar o arquivo: " + nomeArquivoJar);
		            } catch (IOException e) {
		                System.err.println("Erro ao criar o arquivo de erro na subpasta " + subpastaDestino.toString());
		                e.printStackTrace();
		            }

		            // Criar o arquivo de erro na pasta central de erros
		            Path arquivoErroCentral = Paths.get(pastaErros, "Não foi possível baixar - " + pasta + ".txt");
		            try (BufferedWriter writerErro = new BufferedWriter(new FileWriter(arquivoErroCentral.toFile()))) {
		                writerErro.write("Não foi possível baixar o arquivo: " + nomeArquivoJar);
		            } catch (IOException e) {
		                System.err.println("Erro ao criar o arquivo de erro na pasta central " + pastaErros);
		                e.printStackTrace();
		            }
		        }
		    } catch (IOException e) {
		        System.err.println("Erro ao processar o arquivo " + nomeArquivoJar);
		        e.printStackTrace();
		    }
		}
	}
	
	// Percorrer as subpastas e extrair arquivos .jar
	private static void extrairJarFormularioSubpasta(List<String> nomeDasPastas, String pastaFormulariosProd) {
		
		for (String pasta : nomeDasPastas) {
			Path subpastaDestino = Paths.get(pastaFormulariosProd, pasta);
			String nomeArquivoJar = pasta + ".jar";
			Path jarFilePath = subpastaDestino.resolve(nomeArquivoJar);

			if (Files.exists(jarFilePath)) {
				extrairJar(jarFilePath.toString(), subpastaDestino.toString());
			}
			System.out.println("Arquivo Extraido em " + subpastaDestino);
		}
	}

	// Método para extrair o conteúdo de um arquivo .jar
	private static void extrairJar(String caminhoJar, String destinoDiretorio) {
		try (ZipInputStream zis = new ZipInputStream(new FileInputStream(caminhoJar))) {
			ZipEntry zipEntry;
			while ((zipEntry = zis.getNextEntry()) != null) {
				File arquivoDestino = new File(destinoDiretorio, zipEntry.getName());

				if (!zipEntry.isDirectory()) {
					// Cria os diretórios necessários
					new File(arquivoDestino.getParent()).mkdirs();

					// Extrai o arquivo
					try (FileOutputStream fos = new FileOutputStream(arquivoDestino)) {
						byte[] buffer = new byte[1024];
						int bytesRead;
						while ((bytesRead = zis.read(buffer)) != -1) {
							fos.write(buffer, 0, bytesRead);
						}
					}
				}
				zis.closeEntry();
			}
			System.out.println("Arquivos extraídos de " + caminhoJar + " para " + destinoDiretorio);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// Método para copiar arquivos QFS e JASPER para pastas específicas
    private static void copiarArquivosPastasEspecificas(String pastaFormulariosProd, String diretorioProjeto) {
        try {
            // Caminho das pastas para armazenar os arquivos QFS e JASPER
            Path pastaQFS = Paths.get(diretorioProjeto+"/PastaQFS");
            Path pastaJasper = Paths.get(diretorioProjeto+"/PastaJasper");

            // Cria as pastas se não existirem
            Files.createDirectories(pastaQFS);
            Files.createDirectories(pastaJasper);

            // Percorre todas as subpastas de pastaFormulariosProd
            Files.walk(Paths.get(pastaFormulariosProd))
                    .filter(Files::isRegularFile) // Filtra apenas arquivos
                    .forEach(file -> {
                        try {
                            String fileName = file.getFileName().toString().toLowerCase();
                            Path subpastaOrigem = file.getParent().getFileName(); // Nome da subpasta original

                            if (fileName.endsWith(".qfs")) {
                                // Copia o arquivo .QFS para a pastaQFS
                                Files.copy(file, pastaQFS.resolve(file.getFileName()), StandardCopyOption.REPLACE_EXISTING);
                                System.out.println("Arquivo " + file.getFileName() + " copiado para " + pastaQFS);
                            } else if (fileName.endsWith(".jasper")) {
                                // Cria uma subpasta dentro de PastaJasper com o nome da subpasta original
                                Path subpastaJasper = pastaJasper.resolve(subpastaOrigem);
                                Files.createDirectories(subpastaJasper);

                                // Copia o arquivo .JASPER para a subpasta correspondente
                                Files.copy(file, subpastaJasper.resolve(file.getFileName()), StandardCopyOption.REPLACE_EXISTING);
                                System.out.println("Arquivo " + file.getFileName() + " copiado para " + subpastaJasper);
                            }
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    });

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

	// Método para gerar a planilha Excel
	private static void gerarPlanilhaExcel(String pastaFormulariosProd, String diretorioProjeto) {
		String caminhoPlanilha = diretorioProjeto + "/ArquivosFormularios.xlsx";
		List<String> pastasQFS  = new ArrayList<>();
		List<String> pastasJasper = new ArrayList<>();

		// Listar arquivos QFS e Jasper
        try {
            // Listar pastas QFS e Jasper diretamente das pastas PastaQFS e PastaJasper
            Path pastaQFS = Paths.get(diretorioProjeto+"/PastaQFS");
            Path pastaJasper = Paths.get(diretorioProjeto+"/PastaJasper");

            // Listar pastas dentro de PastaQFS
            if (Files.exists(pastaQFS) && Files.isDirectory(pastaQFS)) {
                Files.list(pastaQFS)
                    .forEach(subpasta -> pastasQFS.add(subpasta.getFileName().toString()));
            }

            // Listar pastas dentro de PastaJasper
            if (Files.exists(pastaJasper) && Files.isDirectory(pastaJasper)) {
                Files.list(pastaJasper)
                    .filter(Files::isDirectory)
                    .forEach(subpasta -> pastasJasper.add(subpasta.getFileName().toString()));
            }
            
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Criar a planilha
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Formularios");

            // Cabeçalhos
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(1).setCellValue("Pastas QFS");
            headerRow.createCell(5).setCellValue("Pastas Jasper");

            // Preenchendo a planilha
            int maxRows = Math.max(pastasQFS.size(), pastasJasper.size());
            for (int i = 0; i < maxRows; i++) {
                Row row = sheet.createRow(i + 1);
                if (i < pastasQFS.size()) {
                    row.createCell(1).setCellValue(pastasQFS.get(i));
                }
                if (i < pastasJasper.size()) {
                    row.createCell(5).setCellValue(pastasJasper.get(i) + ".jasper");
                }
            }

            // Salvar a planilha no diretório especificado
            try (FileOutputStream fileOut = new FileOutputStream(caminhoPlanilha)) {
                workbook.write(fileOut);
                System.out.println("Planilha criada com sucesso: " + caminhoPlanilha);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
	
	private static void verificarFormularios(String pastaFormulariosProd, String diretorioBase, String diretorioProjeto) {
        String caminhoPlanilha = diretorioProjeto + "/ArquivosFormularios.xlsx";
        
        try (FileInputStream fis = new FileInputStream(caminhoPlanilha);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Formularios");

            if (sheet == null) {
                System.out.println("A planilha 'Formularios' não foi encontrada.");
                return;
            }

            // Adicionar novo cabeçalho para a nova coluna
            Row headerRow = sheet.getRow(0);
            headerRow.createCell(3).setCellValue("Existe arquivo Jasper ou Jrxml no projeto templates referente aos QFS que estão em produção?");

            // Verificar arquivos listados na pastaQFS e atualizar a planilha
            Path pastaQFS = Paths.get(diretorioProjeto + "/PastaQFS");

            if (Files.exists(pastaQFS) && Files.isDirectory(pastaQFS)) {
                Files.list(pastaQFS)
                    .filter(Files::isRegularFile)
                    .forEach(arquivoQFS -> {
                    	String nomeArquivo = arquivoQFS.getFileName().toString();
                    	String numeroFormulario = nomeArquivo.replaceFirst("\\.QFS$", "");
                        Path pastaCorrespondente = Paths.get(diretorioBase, numeroFormulario);

                        String resultado = "NÃO";
                        if (Files.exists(pastaCorrespondente) && Files.isDirectory(pastaCorrespondente)) {
                            try {
                                boolean encontrouJasperOuJrxml = Files.list(pastaCorrespondente)
                                    .anyMatch(file -> file.toString().endsWith(".jasper") || file.toString().endsWith(".jrxml"));

                                if (encontrouJasperOuJrxml) {
                                    resultado = "SIM";
                                    System.out.println("Arquivos .jasper ou .jrxml encontrados na pasta: " + numeroFormulario);
                                } else {
                                    System.out.println("Nenhum arquivo .jasper ou .jrxml encontrado na pasta: " + numeroFormulario);
                                }
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        } else {
                            System.out.println("Pasta correspondente não encontrada no diretorioBase: " + numeroFormulario);
                        }

                        // Preencher a nova coluna com o resultado
                        int rowIndex = findRowIndex(sheet, numeroFormulario+".QFS", 1); // Usar coluna 1 para "Pastas QFS"
                        if (rowIndex != -1) {
                            Row row = sheet.getRow(rowIndex);
                            Cell novaCelula = row.createCell(3);
                            novaCelula.setCellValue(resultado);
                        }
                        System.out.println("Resultado " + resultado);
                    });
            }

            // Salvar as mudanças na planilha Excel
            try (FileOutputStream fos = new FileOutputStream(caminhoPlanilha)) {
                workbook.write(fos);
                System.out.println("Planilha atualizada com sucesso: " + caminhoPlanilha);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
	}
	
    // Função auxiliar para encontrar o índice da linha que contém o número do formulário
    private static int findRowIndex(Sheet sheet, String numeroFormulario, int colunaIndex) {
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getCell(colunaIndex) != null) {
                if (row.getCell(colunaIndex).getStringCellValue().equals(numeroFormulario)) {
                    return i;
                }
            }
        }
        return -1;
    }
}
