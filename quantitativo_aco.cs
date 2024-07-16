using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using static quantitativosExtraidosTQS.MainWindow;

namespace quantitativosExtraidosTQS
{
    internal class quantitativo_aco
    {
        public static List<Ferro_Resumo> obter_quantitativo_aco(string caminhoPasta)
        {
            Edificio edificio = new Edificio();
            string[] lstFiles = Directory.GetFiles(caminhoPasta, "*.lst", SearchOption.AllDirectories);
            List<Ferro_Resumo> ferros_resumo_projeto = new List<Ferro_Resumo>();
            foreach (string file in lstFiles)
            {
                List<string> caminhoArquivoSeparado = file.Split("\\").ToList();
                if (caminhoArquivoSeparado[^1].EndsWith("_TB.LST") && caminhoArquivoSeparado[^1] != "TABFER_TB.LST") //Usa apenas as tabelas de ferro das folhas, dispensando os das vigas e pilares
                {
                    Pavimento pavimento = new Pavimento();
                    string fileContent = File.ReadAllText(file, Encoding.Latin1);
                    List<string> prePosResumo = fileContent.Split("RESUMO DE AÇO\r\n").ToList();
                    if (prePosResumo.Count() > 1)
                    {
                        List<string> preResumo = prePosResumo[0].Split("\r\n").ToList();
                        List<string> posResumo = prePosResumo[1].Split("\r\n").ToList();
                        List<List<string>> posResumo_limpo = new List<List<string>>();
                        foreach (string item in posResumo)
                        {
                            posResumo_limpo.Add(NormalizarEspacos(item).TrimEnd().TrimStart().Split("  ").ToList());
                        }
                        foreach (List<string> possivel_aco in posResumo_limpo)
                        {
                            if (possivel_aco.Count() == 4 && possivel_aco != posResumo_limpo[0])
                            {
                                Ferro_Resumo ferro_Resumo = new Ferro_Resumo();
                                ferro_Resumo.prancha = caminhoArquivoSeparado[^1].Split(".")[0];
                                ferro_Resumo.nome = possivel_aco[0];
                                ferro_Resumo.bitola = double.Parse(possivel_aco[1], CultureInfo.InvariantCulture);
                                ferro_Resumo.comprimentoTotal = double.Parse(possivel_aco[2], CultureInfo.InvariantCulture);
                                ferro_Resumo.peso = double.Parse(possivel_aco[3], CultureInfo.InvariantCulture);
                                ferros_resumo_projeto.Add(ferro_Resumo);
                            }
                        }
                    }
                   
                }
            }
            return ferros_resumo_projeto;

        }

    }
}
