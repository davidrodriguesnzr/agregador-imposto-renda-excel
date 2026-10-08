# IR Organiza — Agregador de Dados para Imposto de Renda

Projeto prático de Excel desenvolvido para o desafio da **DIO**.

## Sobre o projeto
O **IR Organiza** é uma ferramenta de **organização documental e financeira**, criada para centralizar informações e facilitar a conferência dos dados antes do preenchimento da declaração de Imposto de Renda da Pessoa Física (IRPF). Não gera declaração e não calcula imposto a pagar ou restituição.

A versão compartilhada contém **somente dados fictícios**, com ano-calendário **2025** e exercício de referência **2026**. O ano e o enquadramento fiscal devem ser verificados pelo usuário segundo as regras oficiais pertinentes.

## Funcionalidades
- **Painel:** visão consolidada dos valores informados e controle do checklist.
- **Titular:** identificação fictícia para organização dos dados.
- **Bancos:** instituições, contas mascaradas, saldos e controle dos informes.
- **Rendimentos:** fontes pagadoras, categorias, classificação e valores.
- **Pagamentos:** despesas e evidências documentais.
- **Bens e Direitos:** registros de bens, valores de referência e observações.
- **Dividas:** informações e saldos de obrigações.
- **Checklist:** documentos, fonte, status e acompanhamento de pendências.
- **Listas e Ajuda:** manual de uso e links oficiais.

## Como usar
1. Abra `agregador_imposto_renda_dio.xlsx` no Microsoft Excel.
2. Preencha os campos **azul-claros**. Os dados atualmente são apenas exemplos.
3. Cadastre cada informação em uma linha, conforme sua categoria.
4. Use as listas suspensas para padronizar o preenchimento.
5. Atualize o status dos comprovantes no Checklist.
6. Consulte o Painel e confira os valores com os documentos originais.
7. Utilize os dados conferidos no sistema oficial da Receita Federal.

## Recursos do Excel utilizados
- Fórmulas `SOMA`, `CONT.VALORES` e `CONT.SE` (representadas internamente no XLSX como `SUM`, `COUNTA` e `COUNTIF`).
- Referências a intervalos de outras planilhas e referências absolutas.
- Menus suspensos de validação de dados.
- Formatação condicional para sinalizar pendências e conclusões.
- Estilização de interface com áreas de entrada e indicadores automáticos.
- Congelamento do cabeçalho nos cadastros.
- Gráfico-resumo no dashboard.

Exemplo de fórmula de consolidação:

```excel
=SOMA(Rendimentos!$H$8:$H$107)
```

Exemplo de controle de pendências:

```excel
=CONT.SE(Checklist!$G$8:$G$37;"Pendente")
```

## Estrutura de arquivos
```text
agregador-imposto-renda-excel/
├── README.md
├── agregador_imposto_renda_dio.xlsx
└── images/                 # Capturas de tela opcionais
```

## Demonstração
Os lançamentos ilustrativos resultam em:
- Rendimentos cadastrados: **R$ 42.650,00**
- Pagamentos cadastrados: **R$ 1.250,00**
- Bens cadastrados: **R$ 35.000,00**
- Dívidas cadastradas: **R$ 4.000,00**

Esses valores são apenas **somas de categorias** e não representam base de cálculo, imposto devido ou bens efetivamente sujeitos a declaração.

## Diferenciais do projeto
- Estrutura modular em nove abas.
- Interface que distingue entradas editáveis de resultados automáticos.
- Controle de comprovantes com indicadores.
- Campos e categorias padronizados para facilitar conferências.
- Dados de demonstração anonimizados/fictícios para exposição pública.
- Documentação de limitações e referências à Receita Federal.

## Limitações importantes
Esta ferramenta não avalia obrigatoriedade de declaração, critérios de dedução, tributação específica, classificação final dos bens, dependentes, restituição ou valor de imposto a recolher. Os itens são destinados à organização e devem ser conferidos no material oficial aplicável ao exercício.

**Privacidade:** nunca publique dados verdadeiros, CPF, informações bancárias, recibos médicos nem documentos de terceiros no repositório público.

## Referência oficial
- [Meu Imposto de Renda — Receita Federal](https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda)

## Entrega na DIO
Publique este README e a planilha em um repositório **público** no GitHub. Compartilhe o link do repositório (não apenas do arquivo Excel) na plataforma DIO.

## Contexto
Exercício educacional de modelagem de dados, fórmulas, validação e organização de informações no Excel.
