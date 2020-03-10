using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace ApiAtalho
{
    class Program
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private const string GoogleCredentialsFileName = "google-credentials.json";

        //Planilha Apontamento Wagner
        //private const string idPlanilhaApontamento = "1942C3Hh_V7f9nNzkQ1L-kbtk1KIPg9YTdj86PA7A24s";
        private const string idPlanilhaApontamento = "1YAts-_M4gKFsgJKWqQUiccogtDzDiStWA4joRQvCKOI";

        // Planilha Apontamento Diario - Suporte
        private const string idPlanilhaDiaria = "1z2euyFZOnmuizYtm-wV-L_5HMS_SkZyNJ7O0oJXqFz8";

        private const string lerRangeDiario = "Março/2020!A:D";
        private static string EscreverRangeDiario = "A1:J1";

        private static string[] linhaSplit;
        private static string apontamento;

        static async Task Main(string[] args)
        {
            var serviceValues = GetSheetsService().Spreadsheets.Values;

            Console.WriteLine("Linha: ");
            string linhaPlanilha = Console.ReadLine();

            // Planilha de Apontamento de Horas Wagner
            string lerRangeApontamento = $"Março/2020!A{linhaPlanilha}:J{linhaPlanilha}";

            await LerPlanilha(serviceValues, lerRangeApontamento, idPlanilhaApontamento, false);

            Email email = new Email();
            email.EnviarEmail(linhaSplit, apontamento);
            Console.Read();
        }

        private static SheetsService GetSheetsService()
        {
            using (var stream = new FileStream(GoogleCredentialsFileName,
                FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
                };
                return new SheetsService(serviceInitializer);
            }
        }

        private static async Task LerPlanilha(SpreadsheetsResource.ValuesResource valuesResource,
            string range, string idPlanilha, bool planilhaDiaria)
        {
            var response = await valuesResource.Get(idPlanilha, range).ExecuteAsync();
            var values = response.Values;

            if (planilhaDiaria)
                EscreverRangeDiario = $"A{values.Count + 1}:J{values.Count}";
            else
                await EscreverPlanilhaDiaria(valuesResource, values);

        }

        private static async Task EscreverPlanilhaDiaria(SpreadsheetsResource.ValuesResource valuesResource, IList<IList<object>> values)
        {
            await LerPlanilha(valuesResource, lerRangeDiario, idPlanilhaDiaria, true);

            var linha = string.Join("----", values.First().Select(r => r.ToString()));
            linhaSplit = linha.Split("----");

            string dividir = "N";
            string minutosIndividual;
            string pessoaIndividual;
            string[] minutosPessoa;
            string[] pessoa;

            List<IList<object>> lista = new List<IList<object>>();

            if (linhaSplit[1].Contains("/"))
            {
                Console.WriteLine("Dividir apontamento? (S/N)");
                dividir = Console.ReadLine();

                if (dividir.Equals("S"))
                {
                    pessoa = linhaSplit[1].Split("/");
                    minutosPessoa = new string[pessoa.Length];

                    for (var i = 0; i < pessoa.Length; i++)
                    {
                        Console.WriteLine("Minutos da Pessoa: " + pessoa[i]);
                        minutosPessoa[i] = Console.ReadLine();

                        lista.Add(new List<object> { linhaSplit[6], pessoa[i], linhaSplit[7], int.Parse(minutosPessoa[i]) });
                        apontamento = apontamento + pessoa[i] + " - " + minutosPessoa[i] + " minutos <br/>";
                    }

                    string[] range = EscreverRangeDiario.Split(":");
                    int rangeMax = int.Parse(range[1].Split("J")[1]) + pessoa.Length;
                    EscreverRangeDiario = range[0] + ":J" + rangeMax;
                }
                else
                {
                    Console.WriteLine("Nome: ");
                    pessoaIndividual = Console.ReadLine();
                    Console.WriteLine("Minutos adicionados: ");
                    minutosIndividual = Console.ReadLine();

                    lista.Add(new List<object> { linhaSplit[6], pessoaIndividual, linhaSplit[7], int.Parse(minutosIndividual) });
                    apontamento = pessoaIndividual + " - " + minutosIndividual + " minutos <br/>";
                }
            }
            else { 
                lista.Add(new List<object> { linhaSplit[6], linhaSplit[1], linhaSplit[7], int.Parse(linhaSplit[9]) });
                apontamento = linhaSplit[1] + " - " + linhaSplit[9] + " minutos <br/>";
            }

            if (dividir.Equals("N"))
            {
                string[] range = EscreverRangeDiario.Split(":");
                int rangeMax = int.Parse(range[1].Split("J")[1]) + 1;
                EscreverRangeDiario = range[0] + ":J" + rangeMax;
            }

            var valueRange = new ValueRange
            {
                Values = lista
            };

            var update = valuesResource.Update(valueRange, idPlanilhaDiaria, EscreverRangeDiario);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

            var response = await update.ExecuteAsync();
            Console.WriteLine($"Update row: {response.UpdatedRows}");

        }

    }
}
