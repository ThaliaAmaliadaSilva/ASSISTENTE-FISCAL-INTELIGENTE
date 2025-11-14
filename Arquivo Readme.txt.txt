TÍTULO DO PROJETO:
Assistente Fiscal Inteligente – Análise e Correção Automática de SPED Fiscal

INTEGRANTES:
Thalia Amalia da Silva - 6º Semestre CIC
Ana Luiza Freitas Guimaraes - 6º Semestre CIC

DESCRIÇÃO DO PROCESSO OU PROBLEMA MAPEADO:
O departamento fiscal do escritório de contabilidade realiza diariamente o processo de conferência e correção das notas fiscais de entrada e saída para a geração e entrega do SPED Fiscal (EFD ICMS/IPI). Essa atividade envolve o cruzamento de CFOPs, CSTs, alíquotas e valores, além da análise de erros tributários.

Atualmente, o processo é feito manualmente, em sistemas lentos e com múltiplos parâmetros técnicos. Isso gera sobrecarga de trabalho, alto risco de erro e atraso nas entregas fiscais. O principal gargalo é a falta de automação na etapa de conferência das notas fiscais e na verificação de divergências tributárias.

SOLUÇÃO TECNOLÓGICA PROPOSTA:
Criar um Assistente Fiscal Inteligente que integre o uso de Inteligência Artificial (ChatGPT) e automação em Excel (VBA/Macro). Essa ferramenta tem como objetivo:

1. Ler automaticamente planilhas de notas fiscais (entradas e saídas);
2. Identificar inconsistências tributárias, como:
   - CFOP incompatível com CST;
   - Alíquota incorreta (ex.: produto com 17% lançado com 10%);
   - Tributos não destacados ou destacados incorretamente;
3. Gerar um relatório automático com todas as notas que possuem divergências;
4. Auxiliar o analista fiscal na revisão das regras de tributação conforme o RICMS/MT;
5. Otimizar o tempo de conferência e reduzir riscos de erro antes da entrega do SPED Fiscal.

DESCRIÇÃO DO PROTÓTIPO:
O protótipo foi desenvolvido no ChatGPT, utilizando prompts para gerar códigos VBA aplicados em uma planilha Excel.  
O sistema analisa automaticamente as colunas de CFOP, CST e alíquota das notas fiscais e identifica inconsistências. Quando detectadas, as linhas são destacadas em vermelho e o relatório “Relatorio_Erros” é gerado com todas as divergências encontradas.  
O ChatGPT também é utilizado como apoio técnico para explicar as regras fiscais e gerar o código conforme a legislação de Mato Grosso.

CÓPIA DAS RESPOSTAS DO CHATGPT UTILIZADAS:
As respostas utilizadas estão documentadas no arquivo “Prompts_e_Respostas.txt” dentro deste repositório.

CONCLUSÃO:
O Assistente Fiscal Inteligente mostra como a integração entre IA e automação pode solucionar gargalos reais em escritórios de contabilidade.  
A ferramenta otimiza o trabalho do setor fiscal, reduz retrabalho e aumenta a precisão das informações enviadas no SPED Fiscal, aplicando de forma prática os conceitos aprendidos na disciplina.
