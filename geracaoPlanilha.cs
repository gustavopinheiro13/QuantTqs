using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Documents;
using System.Drawing;
using static quantitativosExtraidosTQS.MainWindow;
using System.Linq;
using System;

namespace quantitativosExtraidosTQS
{
    class linhaPlanilha {
        public string nomePavimento { get; set; }
        public double resumoConcretoTotal { get; set; }
        public double resumoConcretoPilares { get; set; }
        public double resumoConcretoVigas { get; set; }
        public double resumoConcretoLajes { get; set; }
        public double resumoFormaPilar { get; set; }
        public double resumoFormaViga { get; set; }
        public double resumoFormaLajesNervuradas { get; set; }
        public double resumoFormaLajesMacicas { get; set; }
        public double resumoFormaEstruturadaTotal { get; set; }
        public double resumoFormaEstruturadaLajes { get; set; }
        public double resumoAcoCA50CA60 { get; set; } = 0;
        public double resumoAcoCP210 { get; set; } = 0;
        public double resumoAcoCA50CA60Total { get; set; } = 0;
        public double resumoAcoCP210Total { get; set; } = 0;
    }

    class geracaoPlanilha
    {
        protected class ResumoPrancha
        {
            public string Prancha { get; set; }
            public Dictionary<double, double> BitolasPeso { get; set; }

            public ResumoPrancha(string prancha)
            {
                Prancha = prancha;
                BitolasPeso = new Dictionary<double, double>();
            }

            public IEnumerator<KeyValuePair<double, double>> GetEnumerator()
            {
                return BitolasPeso.GetEnumerator();
            }

        }
        public static void SalvarListaEmXLS(Edificio edificio)
        {
            // Cria um novo arquivo XLS
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string caminhoDoArquivo = "quantitativos" + '-' + edificio.cliente + '-' + edificio.nomeEdificio + ".xlsx";

            if (File.Exists(caminhoDoArquivo))
            {
                try
                {
                    File.Delete(caminhoDoArquivo);
                }
                catch (System.IO.IOException)
                {
                    MainWindow.MensagemErro("Sem permissão para editar a planilha, verifique se a planilha está aberta ou que você possui permissão para altera-la e tente novamente. Quantitativo não salvo, saindo", 403);
                    Environment.Exit(403);

                }
            }

            var novoArquivo = new FileInfo(caminhoDoArquivo);
            Stream caminhoListaPadrao;
            caminhoListaPadrao = Assembly.GetExecutingAssembly().GetManifestResourceStream("Quant_Tqs.padrao.quantitativoPadrao.xlsx");
            using (var package = new ExcelPackage(caminhoListaPadrao))
            {
                var cabecalho = package.Workbook.Worksheets["Resumo"];

                cabecalho.Cells[2, 2].Value = edificio.cliente;
                cabecalho.Cells[3, 2].Value = edificio.nomeEdificio;
                cabecalho = package.Workbook.Worksheets["Aco_Pranchado"];
                cabecalho.Cells[2, 2].Value = edificio.cliente;
                cabecalho.Cells[3, 2].Value = edificio.nomeEdificio;
                cabecalho = package.Workbook.Worksheets["Aco_Pranchado_bruto"];
                cabecalho.Cells[2, 2].Value = edificio.cliente;
                cabecalho.Cells[3, 2].Value = edificio.nomeEdificio;
                salvarListaDescritiva(package, edificio);
                salvarListaResumo(package, edificio);
                salvarListaAco(package, edificio);

                // Salva o arquivo
                package.SaveAs(caminhoDoArquivo);
                //package.Save();
            }
        }
        public static void salvarListaResumo(ExcelPackage pacote, Edificio edificio)
        {
            // Adiciona uma nova planilha
            //var planilha = package.Workbook.Worksheets.Add("Entradas");
            bool temProtensao = true, quantitativoPreliminar = true;
            var planilhaResumo = pacote.Workbook.Worksheets["Resumo"];
            List<linhaPlanilha> listaLinhas = new List<linhaPlanilha>();
            foreach (Pavimento pavimento in edificio.listaPavimentos)
            {
                double resumoConcretoTotal = 0, resumoAcoCA50CA60 = 0, resumoAcoCP210 = 0, resumoFormaPilar = 0, resumoFormaViga = 0, resumoFormaEstruturadaTotal = 0, resumoConcretoPilares = 0, resumoConcretoVigas = 0, resumoConcretoLajes = 0, resumoFormaEstruturadaVigas = 0, resumoFormaEstruturadaPilares = 0, resumoFormaEstruturadaLajes = 0, areaEstrutura = 0, resumoFormaLajesMacicas = 0, resumoFormaLajesNervuradas = 0;
                int multiplicador_quantitativos = pavimento.definicoes_piso.Count;
                //planilhaResumo.Cells[linhaAtual, 1].Value = pavimento.nomePlanta;
                foreach (Viga viga in pavimento.listaVigas)
                {
                    resumoConcretoVigas = resumoConcretoVigas + viga.volumeConcreto * multiplicador_quantitativos;
                    resumoFormaEstruturadaVigas = resumoFormaEstruturadaVigas + viga.areaEstruturada * multiplicador_quantitativos;
                    resumoFormaViga = resumoFormaViga + viga.areaFormas * multiplicador_quantitativos;
                }
                foreach (Pilar pilar in pavimento.listaPilares)
                {
                    resumoConcretoPilares = resumoConcretoPilares + pilar.volumeConcreto * multiplicador_quantitativos;
                    resumoFormaEstruturadaPilares = resumoFormaEstruturadaPilares + pilar.areaEstruturada * multiplicador_quantitativos;
                    resumoFormaPilar = resumoFormaPilar + pilar.areaFormas * multiplicador_quantitativos;
                }
                foreach (Laje laje in pavimento.listaLajes)
                {
                    resumoConcretoLajes = resumoConcretoLajes + laje.volumeConcreto * multiplicador_quantitativos;
                    resumoFormaEstruturadaLajes = resumoFormaEstruturadaLajes + laje.areaEstruturada * multiplicador_quantitativos;
                    resumoFormaLajesMacicas = resumoFormaLajesMacicas + laje.areaFormasMacicas * multiplicador_quantitativos;
                    resumoFormaLajesNervuradas = resumoFormaLajesNervuradas + laje.areaFormasNervuradas * multiplicador_quantitativos;
                }
                resumoConcretoTotal = resumoConcretoVigas + resumoConcretoPilares + resumoConcretoLajes;

                if (quantitativoPreliminar)
                {
                    resumoFormaEstruturadaTotal = resumoFormaEstruturadaVigas + resumoFormaEstruturadaPilares + resumoFormaEstruturadaLajes;
                    if (temProtensao)
                    {
                        resumoAcoCP210 = resumoFormaEstruturadaTotal * 2.7;
                    }
                    resumoAcoCA50CA60 = resumoConcretoTotal * 105 - resumoAcoCP210 * 2.5;
                }

                edificio.resumoFormaEstruturadaTotal = edificio.resumoFormaEstruturadaTotal + resumoFormaEstruturadaTotal;
                edificio.resumoConcretoTotal = edificio.resumoConcretoTotal + resumoConcretoTotal;
                edificio.resumoFormaEstruturadaVigas = edificio.resumoFormaEstruturadaVigas + resumoFormaEstruturadaVigas;
                edificio.resumoFormaEstruturadaPilares = edificio.resumoFormaEstruturadaPilares + resumoFormaEstruturadaPilares;
                edificio.resumoFormaEstruturadaLajes = edificio.resumoFormaEstruturadaLajes + resumoFormaEstruturadaLajes;
                edificio.resumoConcretoVigas = edificio.resumoConcretoVigas + resumoConcretoVigas;
                edificio.resumoConcretoPilares = edificio.resumoConcretoPilares + resumoConcretoPilares;
                edificio.resumoConcretoLajes = edificio.resumoConcretoLajes + resumoConcretoLajes;
                edificio.resumoFormaLajesMacicas = edificio.resumoFormaLajesMacicas + resumoFormaLajesMacicas;
                edificio.resumoFormaLajesNervuradas = edificio.resumoFormaLajesNervuradas + resumoFormaLajesNervuradas;
                edificio.resumoFormaPilares = edificio.resumoFormaPilares + resumoFormaPilar;
                edificio.resumoFormaVigas = edificio.resumoFormaVigas + resumoFormaViga;
                edificio.resumoAcoCA50CA60 = edificio.resumoAcoCA50CA60 + resumoAcoCA50CA60;
                edificio.resumoAcoCP210 = edificio.resumoAcoCP210 + resumoAcoCP210;
                linhaPlanilha novalinha = new linhaPlanilha();
                string texto_nome_pavimento_multiplicador = "";
                if (multiplicador_quantitativos > 1)
                {
                    texto_nome_pavimento_multiplicador = " (X "+ multiplicador_quantitativos.ToString() + ")";
                }
                novalinha.nomePavimento = pavimento.nomePlanta + texto_nome_pavimento_multiplicador;
                novalinha.resumoConcretoTotal = resumoConcretoTotal;
                novalinha.resumoConcretoPilares = resumoConcretoPilares;
                novalinha.resumoConcretoVigas = resumoConcretoVigas;
                novalinha.resumoConcretoLajes = resumoConcretoLajes;
                novalinha.resumoFormaPilar = resumoFormaPilar;
                novalinha.resumoFormaViga = resumoFormaViga;
                novalinha.resumoFormaLajesNervuradas = resumoFormaLajesNervuradas;
                novalinha.resumoFormaLajesMacicas = resumoFormaLajesMacicas;
                novalinha.resumoFormaEstruturadaTotal = resumoFormaEstruturadaTotal;
                novalinha.resumoAcoCP210 = resumoAcoCP210;
                novalinha.resumoAcoCA50CA60 = resumoAcoCA50CA60;
                listaLinhas.Add(novalinha);

            }
            //linhaPlanilha novalinhaTotal = new linhaPlanilha();
            //novalinhaTotal.nomePavimento = "TOTAL GERAL:";
            //novalinhaTotal.resumoConcretoTotal = edificio.resumoConcretoTotal;
            //novalinhaTotal.resumoConcretoPilares = edificio.resumoConcretoPilares;
            //novalinhaTotal.resumoConcretoVigas = edificio.resumoConcretoVigas;
            //novalinhaTotal.resumoConcretoLajes = edificio.resumoConcretoLajes;
            //novalinhaTotal.resumoFormaPilar = edificio.resumoFormaPilares;
            //novalinhaTotal.resumoFormaViga = edificio.resumoFormaVigas;
            //novalinhaTotal.resumoFormaEstruturadaLajes = edificio.resumoFormaEstruturadaLajes;
            //novalinhaTotal.resumoFormaLajesNervuradas = edificio.resumoFormaLajesNervuradas;
            //novalinhaTotal.resumoFormaLajesMacicas = edificio.resumoFormaLajesMacicas;
            //novalinhaTotal.resumoFormaEstruturadaTotal = edificio.resumoFormaEstruturadaTotal;
            //novalinhaTotal.resumoAcoCP210 = edificio.resumoAcoCP210;
            //novalinhaTotal.resumoAcoCA50CA60 = edificio.resumoAcoCA50CA60;
            //listaLinhas.Add(novalinhaTotal);
            //linhaInicial += 4;
            //double areaFormaTotal = edificio.resumoFormaLajesMacicas + edificio.resumoFormaPilares + edificio.resumoFormaVigas;
            //planilhaResumo.Cells[linhaAtual, 2].Value = edificio.resumoConcretoTotal / edificio.resumoFormaEstruturadaTotal;
            //planilhaResumo.Cells[linhaAtual, 3].Value = edificio.resumoAcoCA50CA60 / edificio.resumoFormaEstruturadaTotal;
            //planilhaResumo.Cells[linhaAtual, 4].Value = areaFormaTotal / edificio.resumoFormaEstruturadaTotal;
            //planilhaResumo.Cells[linhaAtual, 5].Value = edificio.resumoAcoCP210 / edificio.resumoFormaEstruturadaTotal;
            //planilhaResumo.Cells[linhaAtual, 6].Value = edificio.resumoAcoCA50CA60 / edificio.resumoConcretoTotal;
            //listaLinhas.Reverse();
            int linhaInicial = 6;
            int linhaAtual = linhaInicial;
            bool pintar = true;
            foreach (linhaPlanilha linha in listaLinhas)
            {
                planilhaResumo.Cells[linhaInicial, 1].Value = linha.nomePavimento;
                planilhaResumo.Cells[linhaInicial, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                planilhaResumo.Cells[linhaInicial, 2].Formula = "L" + linhaInicial.ToString() + "*H$" + (linhaAtual + 15).ToString(); //linha.resumoAcoCP210;
                planilhaResumo.Cells[linhaInicial, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 2].Style.Numberformat.Format = "0";
                planilhaResumo.Cells[linhaInicial, 3].Formula = "D" + linhaInicial.ToString() + "*H$" + (linhaAtual + 16).ToString() + "-L$" + (linhaAtual + 16).ToString() + "*B" + linhaInicial.ToString();//linha.resumoAcoCA50CA60;
                planilhaResumo.Cells[linhaInicial, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 3].Style.Numberformat.Format = "0";
                planilhaResumo.Cells[linhaInicial, 4].Value = linha.resumoConcretoTotal;
                planilhaResumo.Cells[linhaInicial, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 5].Value = linha.resumoConcretoPilares;
                planilhaResumo.Cells[linhaInicial, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 6].Value = linha.resumoConcretoVigas;
                planilhaResumo.Cells[linhaInicial, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 7].Value = linha.resumoConcretoLajes;
                planilhaResumo.Cells[linhaInicial, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 8].Value = linha.resumoFormaPilar;
                planilhaResumo.Cells[linhaInicial, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 9].Value = linha.resumoFormaViga;
                planilhaResumo.Cells[linhaInicial, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 10].Value = linha.resumoFormaLajesNervuradas;
                planilhaResumo.Cells[linhaInicial, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 11].Value = linha.resumoFormaLajesMacicas;
                planilhaResumo.Cells[linhaInicial, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 12].Value = linha.resumoFormaEstruturadaTotal;
                planilhaResumo.Cells[linhaInicial, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaResumo.Cells[linhaInicial, 12].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                planilhaResumo.InsertRow(linhaInicial, 1);
                if (pintar)
                {
                    pintar = false;
                    ExcelRange intervaloLinha = planilhaResumo.Cells[linhaInicial, 1, linhaInicial, 12];
                    intervaloLinha.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    intervaloLinha.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 240, 242));
                }
                else { pintar = true; }
                linhaAtual++;
            }
            linhaAtual++;
            for (int numero_coluna = 2; numero_coluna <= 12; numero_coluna++)
            {
                planilhaResumo.Cells[linhaAtual, numero_coluna].Formula = "SUM(" + NumeroParaLetra(numero_coluna) + (linhaInicial + 1).ToString() + ":" + NumeroParaLetra(numero_coluna) + (linhaAtual - 1).ToString() + ")";

            }
            planilhaResumo.DeleteRow(linhaInicial);

            // Define os cabeçalhos das colunas
        }
        public static void salvarListaDescritiva(ExcelPackage pacote, Edificio edificio)
        {
            // Adiciona uma nova planilha
            //var planilha = package.Workbook.Worksheets.Add("Entradas");
            var planilhaDescritiva = pacote.Workbook.Worksheets["Descritivo_Concreto"];

            // Define os cabeçalhos das colunas
            // Preenche as células com os dados das pessoas
            int linhaAtual = 4;
            ExcelRange cells;
            foreach (Pavimento pavimento in edificio.listaPavimentos)
            {
                double  resumoFormaViga = 0, resumoConcretoVigas = 0, resumoFormaEstruturadaVigas = 0;
                planilhaDescritiva.Cells[linhaAtual, 1].Value = pavimento.nomePlanta;
                cells = planilhaDescritiva.Cells[linhaAtual, 1, linhaAtual, 6];
                cells.Merge = true;
                cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                linhaAtual++;
                foreach (Viga viga in pavimento.listaVigas)
                {
                    planilhaDescritiva.Cells[linhaAtual, 1].Value = viga.nome;
                    planilhaDescritiva.Cells[linhaAtual, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 2].Value = viga.areaEstruturada;
                    planilhaDescritiva.Cells[linhaAtual, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 3].Value = viga.areaFormas;
                    planilhaDescritiva.Cells[linhaAtual, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 4].Value = viga.volumeConcreto;
                    planilhaDescritiva.Cells[linhaAtual, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 5].Value = viga.comprimentoLinear;
                    planilhaDescritiva.Cells[linhaAtual, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 6].Value = viga.comprimentoMedioVaos;
                    planilhaDescritiva.Cells[linhaAtual, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    resumoFormaViga = resumoFormaViga + viga.areaFormas;
                    resumoFormaEstruturadaVigas = resumoFormaEstruturadaVigas + viga.areaEstruturada;
                    resumoConcretoVigas = resumoConcretoVigas + viga.volumeConcreto;
                    linhaAtual++;
                }
                planilhaDescritiva.Cells[linhaAtual, 1].Value = "TOTAL:";
                planilhaDescritiva.Cells[linhaAtual, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaDescritiva.Cells[linhaAtual, 2].Value = resumoFormaEstruturadaVigas;
                planilhaDescritiva.Cells[linhaAtual, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaDescritiva.Cells[linhaAtual, 3].Value = resumoFormaViga;
                planilhaDescritiva.Cells[linhaAtual, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaDescritiva.Cells[linhaAtual, 4].Value = resumoConcretoVigas;
                planilhaDescritiva.Cells[linhaAtual, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                linhaAtual++;


            }
            linhaAtual = 4;
            foreach (Pavimento pavimento in edificio.listaPavimentos)
            {
                double resumoFormaPilar = 0, resumoFormaEstruturadaPilares = 0, resumoConcretoPilares = 0;
                planilhaDescritiva.Cells[linhaAtual, 8].Value = pavimento.nomePlanta;
                cells = planilhaDescritiva.Cells[linhaAtual, 8, linhaAtual, 11];
                cells.Merge = true;
                cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                linhaAtual++;
                foreach (Pilar pilar in pavimento.listaPilares)
                {
                    planilhaDescritiva.Cells[linhaAtual, 8].Value = pilar.nome;
                    planilhaDescritiva.Cells[linhaAtual, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 9].Value = pilar.areaEstruturada;
                    planilhaDescritiva.Cells[linhaAtual, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 10].Value = pilar.areaFormas;
                    planilhaDescritiva.Cells[linhaAtual, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 11].Value = pilar.volumeConcreto;
                    planilhaDescritiva.Cells[linhaAtual, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //planilhaDescritiva.Cells[linhaAtual, 5].Value = pilar.volumeTopo;
                    resumoFormaPilar = resumoFormaPilar + pilar.areaFormas;
                    resumoConcretoPilares = resumoConcretoPilares + pilar.volumeConcreto;
                    resumoFormaEstruturadaPilares = resumoFormaEstruturadaPilares + pilar.areaEstruturada;
                    linhaAtual++;
                }
                planilhaDescritiva.Cells[linhaAtual, 8].Value = "TOTAL:";
                planilhaDescritiva.Cells[linhaAtual, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaDescritiva.Cells[linhaAtual, 9].Value = resumoFormaEstruturadaPilares;
                planilhaDescritiva.Cells[linhaAtual, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaDescritiva.Cells[linhaAtual, 10].Value = resumoFormaPilar;
                planilhaDescritiva.Cells[linhaAtual, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaDescritiva.Cells[linhaAtual, 11].Value = resumoConcretoPilares;
                planilhaDescritiva.Cells[linhaAtual, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                linhaAtual++;


            }
            linhaAtual = 4;
            foreach (Pavimento pavimento in edificio.listaPavimentos)
            {
                double resumoAreaEstruturadaLajes = 0, resumoAreaFormasLajes = 0, resumoConcretoLajes = 0;
                planilhaDescritiva.Cells[linhaAtual, 13].Value = pavimento.nomePlanta;
                cells = planilhaDescritiva.Cells[linhaAtual, 13, linhaAtual, 16];
                cells.Merge = true;
                cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                linhaAtual++;
                foreach (Laje laje in pavimento.listaLajes)
                {
                    planilhaDescritiva.Cells[linhaAtual, 13].Value = laje.nome;
                    planilhaDescritiva.Cells[linhaAtual, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 14].Value = laje.areaFormasMacicas;
                    planilhaDescritiva.Cells[linhaAtual, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 15].Value = laje.areaFormasMacicas;
                    planilhaDescritiva.Cells[linhaAtual, 15].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaDescritiva.Cells[linhaAtual, 16].Value = laje.volumeConcreto;
                    planilhaDescritiva.Cells[linhaAtual, 16].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    resumoAreaFormasLajes = resumoAreaFormasLajes + laje.areaFormasMacicas;
                    resumoAreaEstruturadaLajes = resumoAreaEstruturadaLajes + laje.areaEstruturada;
                    resumoConcretoLajes = resumoConcretoLajes + laje.volumeConcreto;
                    linhaAtual++;
                }
                planilhaDescritiva.Cells[linhaAtual, 13].Value = "TOTAL:";
                planilhaDescritiva.Cells[linhaAtual, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaDescritiva.Cells[linhaAtual, 14].Value = resumoAreaEstruturadaLajes;
                planilhaDescritiva.Cells[linhaAtual, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaDescritiva.Cells[linhaAtual, 15].Value = resumoAreaFormasLajes;
                planilhaDescritiva.Cells[linhaAtual, 15].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaDescritiva.Cells[linhaAtual, 16].Value = resumoConcretoLajes;
                planilhaDescritiva.Cells[linhaAtual, 16].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                linhaAtual++;
            }
        }
        public static void salvarListaAco(ExcelPackage pacote, Edificio edificio)
        {
            // Adiciona uma nova planilha
            //var planilha = package.Workbook.Worksheets.Add("Entradas");
            bool temProtensao = true, quantitativoPreliminar = true;
            var planilhaAco = pacote.Workbook.Worksheets["Aco_Pranchado_bruto"];
            int linhaInicial = 5;
            int linhaAtual = linhaInicial;
            List<linhaPlanilha> listaLinhas = new List<linhaPlanilha>();
            listaLinhas.Reverse();
            bool pintar = true;
            var bitolasDistintas = edificio.ferro_resumo_pranchado.Select(f => f.bitola).Distinct().OrderBy(b => b);
            int numero_bitolas_distintas = bitolasDistintas.ToList().Count();

            foreach (Ferro_Resumo ferro_resumo_pranchado in edificio.ferro_resumo_pranchado)
            {
                planilhaAco.Cells[linhaAtual, 1].Value = ferro_resumo_pranchado.prancha;
                planilhaAco.Cells[linhaAtual, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaAco.Cells[linhaAtual, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                planilhaAco.Cells[linhaAtual, 2].Value = ferro_resumo_pranchado.nome;
                planilhaAco.Cells[linhaAtual, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaAco.Cells[linhaAtual, 3].Value = ferro_resumo_pranchado.bitola;
                planilhaAco.Cells[linhaAtual, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaAco.Cells[linhaAtual, 3].Style.Numberformat.Format = "0.00";
                planilhaAco.Cells[linhaAtual, 4].Value = ferro_resumo_pranchado.comprimentoTotal;
                planilhaAco.Cells[linhaAtual, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaAco.Cells[linhaAtual, 4].Style.Numberformat.Format = "0.00";
                planilhaAco.Cells[linhaAtual, 5].Value = ferro_resumo_pranchado.peso;
                planilhaAco.Cells[linhaAtual, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaAco.Cells[linhaAtual, 5].Style.Numberformat.Format = "0.00";
                planilhaAco.Cells[linhaAtual, 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                planilhaAco.InsertRow(linhaAtual, 1);
                if (pintar)
                {
                    pintar = false;
                    ExcelRange intervaloLinha = planilhaAco.Cells[linhaAtual, 1, linhaAtual, 5];
                    intervaloLinha.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    intervaloLinha.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 240, 242));
                }
                else { pintar = true; }
            }
            linhaAtual += edificio.ferro_resumo_pranchado.Count();
            linhaAtual++;
            for (int numero_coluna = 4; numero_coluna <= 5; numero_coluna++)
            {
                planilhaAco.Cells[linhaAtual, numero_coluna].Formula = "SUM(" + NumeroParaLetra(numero_coluna) + (linhaInicial + 1).ToString() + ":" + NumeroParaLetra(numero_coluna) + (linhaAtual - 1).ToString() + ")";
                planilhaAco.Cells[linhaAtual, numero_coluna].Style.Numberformat.Format = "0.00";

            }

            planilhaAco.DeleteRow(linhaInicial);

            planilhaAco = pacote.Workbook.Worksheets["Aco_Pranchado"];
            int coluna_atual = 3;
            linhaAtual = linhaInicial;
            for (int i = 0; i < numero_bitolas_distintas -1; i++)
            {
                planilhaAco.InsertColumn(coluna_atual, 1);
            }
            // Agrupa por prancha e bitola e soma os pesos
            var resumoAgrupado = edificio.ferro_resumo_pranchado
                .GroupBy(f => new { f.prancha, f.bitola })
                .Select(g => new { g.Key.prancha, g.Key.bitola, PesoTotal = g.Sum(f => f.peso) })
                .ToList();


            // Cria uma lista de ResumoPrancha para cada prancha com as bitolas e os pesos somados
            var listaResumoPrancha = new List<ResumoPrancha>();

            foreach (var prancha in resumoAgrupado.GroupBy(f => f.prancha))
            {
                var resumoPrancha = new ResumoPrancha(prancha.Key);
                foreach (var bitola in bitolasDistintas)
                {
                    var pesoTotal = prancha.Where(f => f.bitola == bitola).Sum(f => f.PesoTotal);
                    resumoPrancha.BitolasPeso.Add(bitola, pesoTotal);
                }
                listaResumoPrancha.Add(resumoPrancha);
            }
            pintar = true;
            int coluna_inicial = 1;
            coluna_atual = coluna_inicial;
            foreach (double bitola_distinta in bitolasDistintas)
            {
                coluna_atual++;
                planilhaAco.Cells[linhaInicial, coluna_atual].Value = bitola_distinta;
                planilhaAco.Cells[linhaInicial, coluna_atual].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaAco.Cells[linhaInicial, coluna_atual].Style.Numberformat.Format = "0.00";
            }
            for (int i = 0; i <= numero_bitolas_distintas +1; i++)
            {
                for (int j = 1; j < linhaInicial + 1; j++)
                {
                    planilhaAco.Cells[j, coluna_inicial + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    planilhaAco.Cells[j, coluna_inicial + i].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                }
            }
            linhaAtual++;
            foreach (ResumoPrancha resumo_formatado in listaResumoPrancha)
            {
                coluna_atual = coluna_inicial;
                planilhaAco.Cells[linhaInicial + 1, coluna_atual].Value = resumo_formatado.Prancha;
                planilhaAco.Cells[linhaInicial + 1, coluna_atual].Style.Border.Left.Style = ExcelBorderStyle.Thin;

                coluna_atual++;
                double pesoTotal = 0;
                foreach (KeyValuePair<double,double> bitola_peso in resumo_formatado)
                {
                    planilhaAco.Cells[linhaInicial + 1, coluna_atual].Value = bitola_peso.Value;
                    planilhaAco.Cells[linhaInicial + 1, coluna_atual].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    planilhaAco.Cells[linhaInicial + 1, coluna_atual].Style.Numberformat.Format = "0.00";
                    coluna_atual++;
                    pesoTotal += bitola_peso.Value;
                }
                planilhaAco.Cells[linhaInicial + 1, coluna_atual].Formula = "SUM(" + NumeroParaLetra(coluna_inicial +1) + (linhaInicial + 1).ToString() + ":" + NumeroParaLetra(coluna_atual-1) + (linhaInicial + 1).ToString() + ")";
                planilhaAco.Cells[linhaInicial + 1, coluna_atual].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaAco.Cells[linhaInicial + 1, coluna_atual].Style.Numberformat.Format = "0.00";
                planilhaAco.Cells[linhaInicial + 1, coluna_atual].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                if (pintar)
                {
                    pintar = false;
                    ExcelRange intervaloLinha = planilhaAco.Cells[linhaInicial + 1, 1, linhaInicial + 1, numero_bitolas_distintas + 2];
                    intervaloLinha.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    intervaloLinha.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 240, 242));
                }
                else { pintar = true; }
                planilhaAco.InsertRow(linhaInicial + 1, 1);
                linhaAtual++;

            }
            planilhaAco.DeleteRow(linhaInicial + 1);
            //linhaAtual++;
            for (int i = 0; i <= numero_bitolas_distintas; i++)
            {
                planilhaAco.Cells[linhaAtual, coluna_inicial + 1 + i].Formula = "SUM(" + NumeroParaLetra(coluna_inicial + 1 + i) + (linhaInicial +1).ToString() + ":" + NumeroParaLetra(coluna_inicial + 1 + i) + (linhaAtual - 1).ToString() + ")";//linha.resumoAcoCA50CA60;;
                planilhaAco.Cells[linhaInicial -1, coluna_inicial + 1 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                planilhaAco.Cells[linhaInicial -1, coluna_inicial + 1 + i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(9, 230, 176));
                planilhaAco.Cells[linhaAtual, coluna_inicial + 1 + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                planilhaAco.Cells[linhaAtual, coluna_inicial + 1 + i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(9, 230, 176));
                planilhaAco.Cells[linhaAtual, coluna_inicial + 1 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                planilhaAco.Cells[linhaAtual, coluna_inicial + 1 + i].Style.Numberformat.Format = "0.00";
                planilhaAco.Cells[linhaAtual, coluna_inicial + 1 + i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                planilhaAco.Cells[linhaAtual, coluna_inicial + 1 + i].Style.Border.Top.Style = ExcelBorderStyle.Thin;

            }
            //planilhaAco.Cells[linhaAtual, coluna_inicial + 2 + numero_bitolas_distintas].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            // Define os cabeçalhos das colunas
        }
        public static string NumeroParaLetra(int numero)
        {
            if (numero <= 0)
            {
                return "";
            }

            string resultado = "";
            while (numero > 0)
            {
                numero--; // Decrementa para mapear 1 a 'A', 2 a 'B', etc.
                char letra = (char)('A' + (numero % 26));
                resultado = letra + resultado;
                numero /= 26;
            }

            return resultado;
        }

    }
}
