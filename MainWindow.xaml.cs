using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System;
using System.IO;
using System.Collections;
using System.Globalization;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.ComponentModel;
using System.Reflection;
using static quantitativosExtraidosTQS.MainWindow;
using System.IO.Packaging;
using Microsoft.Win32;
using System.Xml;
using System.Text.RegularExpressions;
using Microsoft.WindowsAPICodePack.Dialogs;
using quantitativosExtraidosTQS;

namespace quantitativosExtraidosTQS
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public class Edificio
        {
            public string nomeEdificio { get; set; } = "";
            public string tituloGeral { get; set; } = "";
            public string cliente { get; set; } = "";
            public double resumoConcretoTotal { get; set; } = 0;
            public double resumoFormaEstruturadaTotal { get; set; } = 0;
            public double resumoFormaVigas { get; set; } = 0;
            public double resumoFormaPilares { get; set; } = 0;
            public double resumoConcretoPilares { get; set; } = 0;
            public double resumoConcretoVigas { get; set; } = 0;
            public double resumoConcretoLajes { get; set; } = 0;
            public double resumoFormaEstruturadaPilares { get; set; } = 0;
            public double resumoFormaEstruturadaVigas { get; set; } = 0;
            public double resumoFormaEstruturadaLajes { get; set; } = 0;
            public double resumoFormaLajesMacicas { get; set; } = 0;
            public double resumoFormaLajesNervuradas { get; set; } = 0;
            public double resumoAcoCA50CA60 { get; set; } = 0;
            public double resumoAcoCP210 { get; set; } = 0;
            public List<Pavimento> listaPavimentos { get; set; } = new List<Pavimento>();
            public List<Ferro_Resumo> ferro_resumo_pranchado { get; set; } = new List<Ferro_Resumo>();
        }
        public class Pavimento
        {
            public string nomeEdificio { get; set; } = "";
            public string nomePlanta { get; set; } = "";
            public string tituloGeral { get; set; } = "";
            public string cliente { get; set; } = "";
            public string tituloPlanta { get; set; } = "";
            public List<Definicao_Piso> definicoes_piso { get; set; } = new List<Definicao_Piso>();
            public List<Viga> listaVigas { get; set; } = new List<Viga>();
            public List<Pilar> listaPilares { get; set; } = new List<Pilar>();
            public List<Laje> listaLajes { get; set; } = new List<Laje>();
        }
        public class Viga
        {
            public string nome { get; set; } = "";
            public double areaEstruturada { get; set; } = 0;
            public double areaFormas { get; set; } = 0;
            public double volumeConcreto { get; set; } = 0;
            public double comprimentoLinear { get; set; } = 0;
            public double comprimentoMedioVaos { get; set; } = 0;
            //public List<Ferro> armaduras { get; set; } = new List<Ferro>();

        }
        public class Laje
        {
            public string nome { get; set; } = "";
            public double areaEstruturada { get; set; } = 0;
            public double areaFormasMacicas { get; set; } = 0;
            public double areaFormasNervuradas { get; set; } = 0;
            public double volumeConcreto { get; set; } = 0;
            //public List<Ferro> armaduras { get; set; } = new List<Ferro>();

        }
        public class Pilar
        {
            public string nome { get; set; }
            public double areaEstruturada { get; set; }
            public double areaFormas { get; set; }
            public double volumeConcreto { get; set; }
            //public double volumeTopo { get; set; }
            //public List<Ferro> armaduras { get; set; } = new List<Ferro>();

        }
        public class Ferro
        {
            public string nome { get; set; }
            public int posicao { get; set; }
            public double bitola { get; set; }
            public int quantidade { get; set; }
            public double comprimentoUnitario { get; set; }
            public double comprimentoTotal { get; set; }
        }
        public class Ferro_Resumo
        {
            public string prancha { get; set; }
            public string nome { get; set; }
            public double bitola { get; set; }
            public double comprimentoTotal { get; set; }
            public double peso { get; set; }

        }
        public class Definicao_Piso
        {
            public string nome { get; set; }
            public int numero_piso { get; set; }
            public double cota { get; set; }
            public double pe_direito { get; set; }
            public double secao { get; set; }
        }
        public static List<Ferro> armadurasEdificio { get; set; } = new List<Ferro>();
        public static string RemoverEspacosRepetidos(string texto)
        {
            return Regex.Replace(texto, @"\s+", " ");
        }
        public static string NormalizarEspacos(string texto)
        {
            return Regex.Replace(texto, @"\s{3,}", "  ");
        }

        static Edificio gerarEdificio(string caminhoPasta)
        {
            Edificio edificio = new Edificio();
            string[] lstFiles = Directory.GetFiles(caminhoPasta, "*.lst", SearchOption.AllDirectories);

            foreach (string file in lstFiles)
            {
                List<string> caminhoArquivoSeparado = file.Split("\\").ToList();
                if (caminhoArquivoSeparado[caminhoArquivoSeparado.Count - 2] + ".LST" == caminhoArquivoSeparado[caminhoArquivoSeparado.Count - 1])
                {
                    Pavimento pavimento = new Pavimento();
                    //Console.WriteLine($"Arquivo: {file}");
                    string fileContent = File.ReadAllText(file, Encoding.Latin1);
                    List<string> prePosQuantitativos = fileContent.Split("Quantitativos\r\n-------------").ToList();
                    List<string> preQuantitativos = prePosQuantitativos[0].Split("\r\n").ToList();
                    int intidiceCabecalho = preQuantitativos.IndexOf("=======================");
                    pavimento.nomeEdificio = preQuantitativos[intidiceCabecalho + 1].Split(". ").Last<string>();
                    pavimento.nomePlanta = preQuantitativos[intidiceCabecalho + 3].Split(". ").Last<string>();
                    pavimento.tituloGeral = preQuantitativos[intidiceCabecalho + 5].Split(". ").Last<string>();
                    pavimento.cliente = preQuantitativos[intidiceCabecalho + 6].Split(". ").Last<string>();
                    pavimento.tituloPlanta = preQuantitativos[intidiceCabecalho + 7].Split(". ").Last<string>();
                    //int.Parse(preQuantitativos[intidiceCabecalho + 16].Split(" ")[1])
                    List<int> numeros_repeticao_planta = new List<int>();
                    List<string> secao_numero_pavimentos = prePosQuantitativos[0].Split("Definição de Pisos\r\n------------------")[1].Split("\r\n").ToList();
                    List<Definicao_Piso> definicoes_pavimento = new List<Definicao_Piso>();
                    foreach (string linha in secao_numero_pavimentos)
                    {

                        List<string> linha_dividida = NormalizarEspacos(linha).TrimEnd().TrimStart().Split("  ").ToList();
                        if (linha_dividida.Count() == 6 && int.TryParse(linha_dividida[0], out int numero) || linha_dividida.Count() == 7 && int.TryParse(linha_dividida[0], out int numero2) || linha_dividida.Count() == 5 && int.TryParse(linha_dividida[0], out int numero3))
                        {
                            linha_dividida.RemoveAll(string.IsNullOrEmpty);
                            List<string> definicao_linha_limpa = new List<string>();
                            foreach (string item in linha_dividida)
                            {
                                definicao_linha_limpa.Add(item);
                            }
                            Definicao_Piso nova_definicao = new Definicao_Piso();
                            nova_definicao.nome = definicao_linha_limpa[1];
                            nova_definicao.numero_piso = int.Parse(definicao_linha_limpa[0]);
                            nova_definicao.secao = int.Parse(definicao_linha_limpa[4]);
                            nova_definicao.cota = double.Parse(definicao_linha_limpa[2], CultureInfo.InvariantCulture);
                            nova_definicao.pe_direito = double.Parse(definicao_linha_limpa[3], CultureInfo.InvariantCulture);
                            definicoes_pavimento.Add(nova_definicao);
                        }
                    }

                    pavimento.definicoes_piso = definicoes_pavimento;

                    List<string> posQuantitativos = prePosQuantitativos[1].Split("\r\n").ToList();
                    List<string> intervaloQuantitativos = new List<string>();
                    List<Viga> listaVigas = new List<Viga>();
                    List<Pilar> listaPilares = new List<Pilar>();
                    List<Laje> listaLajes = new List<Laje>();

                    foreach (string line in posQuantitativos)
                    {
                        intervaloQuantitativos.Add(line);
                        if (line.StartsWith("Espessura"))
                        {
                            break;
                        }
                    }
                    List<string> elementosEstruturais = new List<string>();
                    foreach (string linha in intervaloQuantitativos)
                    {
                        List<string> linhaAtualProcessada = linha.Split(' ').ToList<string>();
                        linhaAtualProcessada.RemoveAll(item => item == "");
                        if (linhaAtualProcessada.Count == 6 && linhaAtualProcessada[0].ToString().Contains("V"))
                        {
                            Viga novaviga = new Viga();
                            novaviga.nome = linhaAtualProcessada[0];
                            novaviga.areaEstruturada = double.Parse('0' + linhaAtualProcessada[1].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaviga.areaFormas = double.Parse('0' + linhaAtualProcessada[2].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaviga.volumeConcreto = double.Parse('0' + linhaAtualProcessada[3].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaviga.comprimentoLinear = double.Parse('0' + linhaAtualProcessada[4].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaviga.comprimentoMedioVaos = double.Parse('0' + linhaAtualProcessada[5].Replace("-", ""), CultureInfo.InvariantCulture);
                            listaVigas.Add(novaviga);
                        }
                        else if (linhaAtualProcessada.Count == 6 && linhaAtualProcessada[0].ToString().Contains("ABA"))
                        {
                            Viga novaviga = new Viga();
                            novaviga.nome = linhaAtualProcessada[0];
                            novaviga.areaEstruturada = double.Parse('0' + linhaAtualProcessada[1].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaviga.areaFormas = double.Parse('0' + linhaAtualProcessada[2].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaviga.volumeConcreto = double.Parse('0' + linhaAtualProcessada[3].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaviga.comprimentoLinear = double.Parse('0' + linhaAtualProcessada[4].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaviga.comprimentoMedioVaos = double.Parse('0' + linhaAtualProcessada[5].Replace("-", ""), CultureInfo.InvariantCulture);
                            listaVigas.Add(novaviga);
                        }
                        else if (linhaAtualProcessada.Count == 5 && linhaAtualProcessada[0].ToString().Contains("P"))
                        {
                            Pilar novaPilar = new Pilar();
                            novaPilar.nome = linhaAtualProcessada[0];
                            novaPilar.areaEstruturada = double.Parse('0' + linhaAtualProcessada[1].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaPilar.areaFormas = double.Parse('0' + linhaAtualProcessada[2].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaPilar.volumeConcreto = double.Parse('0' + linhaAtualProcessada[3].Replace("-", ""), CultureInfo.InvariantCulture);
                            //novaPilar.volumeTopo = double.Parse('0' + linhaAtualProcessada[4].Replace("-", ""), CultureInfo.InvariantCulture);
                            listaPilares.Add(novaPilar);
                        }
                        else if (linhaAtualProcessada.Count == 10 && linhaAtualProcessada[0].ToString().Contains("L"))
                        {
                            Laje novaLaje = new Laje();
                            novaLaje.nome = linhaAtualProcessada[0];
                            novaLaje.areaEstruturada = double.Parse('0' + linhaAtualProcessada[1].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.areaFormasMacicas = double.Parse('0' + linhaAtualProcessada[2].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.volumeConcreto = double.Parse('0' + linhaAtualProcessada[3].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.areaFormasNervuradas = novaLaje.areaEstruturada - novaLaje.areaFormasMacicas;
                            listaLajes.Add(novaLaje);
                        }
                        else if (linhaAtualProcessada.Count == 4 && linhaAtualProcessada[0].ToString().Contains("E"))
                        {
                            Laje novaLaje = new Laje();
                            novaLaje.nome = linhaAtualProcessada[0];
                            novaLaje.areaEstruturada = double.Parse('0' + linhaAtualProcessada[1].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.areaFormasMacicas = double.Parse('0' + linhaAtualProcessada[2].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.volumeConcreto = double.Parse('0' + linhaAtualProcessada[3].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.areaFormasNervuradas = novaLaje.areaEstruturada - novaLaje.areaFormasMacicas;
                            listaLajes.Add(novaLaje);
                        }
                        else if (linhaAtualProcessada.Count == 4 && linhaAtualProcessada[0].ToString().Contains("R"))
                        {
                            Laje novaLaje = new Laje();
                            novaLaje.nome = linhaAtualProcessada[0];
                            novaLaje.areaEstruturada = double.Parse('0' + linhaAtualProcessada[1].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.areaFormasMacicas = double.Parse('0' + linhaAtualProcessada[2].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.volumeConcreto = double.Parse('0' + linhaAtualProcessada[3].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.areaFormasNervuradas = novaLaje.areaEstruturada - novaLaje.areaFormasMacicas;
                            listaLajes.Add(novaLaje);
                        }
                        else if (linhaAtualProcessada.Count == 4 && linhaAtualProcessada[0].ToString().Contains("L"))
                        {
                            Laje novaLaje = new Laje();
                            novaLaje.nome = linhaAtualProcessada[0];
                            novaLaje.areaEstruturada = double.Parse('0' + linhaAtualProcessada[1].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.areaFormasMacicas = double.Parse('0' + linhaAtualProcessada[2].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.volumeConcreto = double.Parse('0' + linhaAtualProcessada[3].Replace("-", ""), CultureInfo.InvariantCulture);
                            novaLaje.areaFormasNervuradas = novaLaje.areaEstruturada - novaLaje.areaFormasMacicas;
                            listaLajes.Add(novaLaje);
                        }
                    }
                    pavimento.listaVigas = listaVigas;
                    pavimento.listaPilares = listaPilares;
                    pavimento.listaLajes = listaLajes;
                    edificio.listaPavimentos.Add(pavimento);
                    edificio.cliente = pavimento.cliente;
                    edificio.tituloGeral = pavimento.tituloGeral;
                    edificio.nomeEdificio = pavimento.nomeEdificio;
                }
            }
            edificio.listaPavimentos = edificio.listaPavimentos.OrderBy(p => p.definicoes_piso.Min(dp => dp.numero_piso)).ToList();
            edificio.ferro_resumo_pranchado = quantitativo_aco.obter_quantitativo_aco(caminhoPasta);
            return edificio;

        }
        public MainWindow()
        {

            //Console.WriteLine("Digite o caminho para a pasta:");
            string folderPath;
            folderPath = selecionarPasta();
            if (folderPath == null)
            {
                MensagemErro("Pasta do projeto não selecionada. Selecione a pasta raiz do Projeto", 415);
                Environment.Exit(415);

            }
            else
            {
                Edificio novoEdificio = gerarEdificio(folderPath);

                if (novoEdificio != null)
                {
                    geracaoPlanilha.SalvarListaEmXLS(novoEdificio);
                    mensagemSucesso();

                }

                //InitializeComponent();
            }
            Application.Current.Shutdown();
        }
        public void mensagemSucesso()
        {
            MessageBox.Show("Sucesso, Quantitativos salvos!", "Concluido", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        public static void MensagemErro(string mensagem_erro, int codigo_erro)
        {
            MessageBox.Show(mensagem_erro, "Erro " + codigo_erro.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
        }
        static string selecionarPasta()
        {
            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "Selecione a pasta do projeto"
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string pastaSelecionada = dialog.FileName;
                string arquivo = System.IO.Path.Combine(pastaSelecionada, "EDIFICIO.BDE");

                if (File.Exists(arquivo))
                {
                    return pastaSelecionada;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                // Caso o usuário cancele a seleção da pasta
                return null;
            }
        }

        }
    }
public class Program
{
    [STAThread]
    public static void Main()
    {
        Application app = new Application();
        MainWindow mainWindow = new MainWindow();
        app.Run(mainWindow);


    }
}