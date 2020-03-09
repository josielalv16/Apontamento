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

namespace ApiAtalho {
	class Program {
		private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
		private const string GoogleCredentialsFileName = "google-credentials.json";

		//Planilha Apontamento Wagner
		private const string idPlanilhaApontamento = "1942C3Hh_V7f9nNzkQ1L-kbtk1KIPg9YTdj86PA7A24s";
		// Planilha Apontamento Diario - Suporte
		private const string idPlanilhaDiaria = "1z2euyFZOnmuizYtm-wV-L_5HMS_SkZyNJ7O0oJXqFz8";

		private const string lerRangeDiario = "Março/2020!A:D";
		private static string EscreverRangeDiario = "A1:J1";

		static async Task Main(string[] args) {
			var serviceValues = GetSheetsService().Spreadsheets.Values;

			Console.WriteLine("Linha: ");
			string linhaPlanilha = Console.ReadLine();

			// Planilha de Apontamento de Horas Wagner
			string lerRangeApontamento = $"Pagina1!A{linhaPlanilha}:J{linhaPlanilha}";

			await LerPlanilha(serviceValues, lerRangeApontamento, idPlanilhaApontamento, false);

			//Console.Read();
		}

		private static SheetsService GetSheetsService() {
			using (var stream = new FileStream(GoogleCredentialsFileName,
				FileMode.Open, FileAccess.Read)) {
				var serviceInitializer = new BaseClientService.Initializer {
					HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
				};
				return new SheetsService(serviceInitializer);
			}
		}

		private static async Task LerPlanilha(SpreadsheetsResource.ValuesResource valuesResource,
			string range, string idPlanilha, bool planilhaDiaria) {
			var response = await valuesResource.Get(idPlanilha, range).ExecuteAsync();
			var values = response.Values;

			if (planilhaDiaria)
				EscreverRangeDiario = $"A{values.Count + 1}:J{values.Count}";
			else
				await EscreverPlanilhaDiaria(valuesResource, values);

		}

		private static async Task EscreverPlanilhaDiaria(SpreadsheetsResource.ValuesResource valuesResource, IList<IList<object>> values) {
			await LerPlanilha(valuesResource, lerRangeDiario, idPlanilhaDiaria, true);

			var linha = string.Join("----", values.First().Select(r => r.ToString()));
			string[] linhaSplit = linha.Split("----");

			string dividir = "N";
			string minutosIndividual;
			string pessoaIndividual;
			string[] minutosPessoa;
			string[] pessoa;

			List<IList<object>> lista = new List<IList<object>>();

			if (linhaSplit[1].Contains("/")) {
				Console.WriteLine("Dividir apontamento? (S/N)");
				dividir = Console.ReadLine();

				if (dividir.Equals("S")) {
					pessoa = linhaSplit[1].Split("/");
					minutosPessoa = new string[pessoa.Length];

					for (var i = 0; i < pessoa.Length; i++) {
						Console.WriteLine("Minutos da Pessoa: " + pessoa[i]);
						minutosPessoa[i] = Console.ReadLine();

						lista.Add(new List<object> { linhaSplit[6], pessoa[i], linhaSplit[7], minutosPessoa[i] });
					}

					string[] range = EscreverRangeDiario.Split(":");
					int rangeMax = int.Parse(range[1].Split("J")[1]) + pessoa.Length;
					EscreverRangeDiario = range[0] + ":J" + rangeMax;
				}
				else {
					Console.WriteLine("Nome: ");
					pessoaIndividual = Console.ReadLine();
					Console.WriteLine("Minutos adicionados: ");
					minutosIndividual = Console.ReadLine();

					lista.Add(new List<object> { linhaSplit[6], pessoaIndividual, linhaSplit[7], minutosIndividual });
				}
			}
			else
				lista.Add(new List<object> { linhaSplit[6], linhaSplit[1], linhaSplit[7], linhaSplit[9] });

			if (dividir.Equals("N")) {
				string[] range = EscreverRangeDiario.Split(":");
				int rangeMax = int.Parse(range[1].Split("J")[1]) + 1;
				EscreverRangeDiario = range[0] + ":J" + rangeMax;
			}

			var valueRange = new ValueRange {
				Values = lista
			};

			var update = valuesResource.Update(valueRange, idPlanilhaDiaria, EscreverRangeDiario);
			update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

			var response = await update.ExecuteAsync();
			Console.WriteLine($"Update row: {response.UpdatedRows}");

			EnviarEmailApontamento(linhaSplit);
		}

		private static void EnviarEmailApontamento(string[] linhaSplit) {
			var body = File.ReadAllText($@"{AppDomain.CurrentDomain.BaseDirectory}modelo.html");
			body = body.Replace("{{nome}}", linhaSplit[1])
						.Replace("{{tempoGasto}}", linhaSplit[9])
						.Replace("{{dataSol}}", linhaSplit[0])
						.Replace("{{analista}}", linhaSplit[1])
						.Replace("{{tipoAtendimento}}", linhaSplit[2])
						.Replace("{{solicitante}}", linhaSplit[3])
						.Replace("{{area}}", linhaSplit[4])
						.Replace("{{dataInicio}}", linhaSplit[5])
						.Replace("{{dataFim}}", linhaSplit[6])
						.Replace("{{assunto}}", linhaSplit[7])
						.Replace("{{descricao}}", linhaSplit[8])
						.Replace("{{tempoGasto}}", linhaSplit[9]);

			var email = new MailMessage("josiel.alves@smn.com.br", "josielalv@hotmail.com") {
				Subject = linhaSplit[7],
				Body = body,
				IsBodyHtml = true,
				Priority = MailPriority.Normal,
				SubjectEncoding = Encoding.GetEncoding("ISO-8859-1"),
				BodyEncoding = Encoding.GetEncoding("ISO-8859-1")
			};

			try {
				using (var smtp = new SmtpClient("smtp-mail.outlook.com", 587) { Credentials = new NetworkCredential("josiel.alves@smn.com.br", "SMN@2017#o365") }) {
					smtp.EnableSsl = true;
					smtp.Send(email);
				}
				Console.WriteLine("E-mail enviado com sucesso!");
			}
			catch (Exception erro) {
				Console.WriteLine("Erro ao enviar E-mail \n\n" + erro);
			}
			finally {
				email = null;
			}
		}

		//private static async Task ReadAsync(SpreadsheetsResource.ValuesResource valuesResource) {
		//	var response = await valuesResource.Get(SpreadsheetId, ReadRange).ExecuteAsync();
		//	var values = response.Values;

		//	if (values == null || !values.Any()) {
		//		Console.WriteLine("No data found.");
		//		return;
		//	}

		//	var header = string.Join(" ", values.First().Select(r => r.ToString()));
		//	Console.WriteLine($"Header: {header}");
		//	foreach (var row in values.Skip(1)) {
		//		var res = string.Join(" ", row.Select(r => r.ToString()));
		//		Console.WriteLine(res);
		//	}

		//	EscreverRange = $"A{values.Count + 1}:J{values.Count + 1}";
		//}

		//private static async Task WriteAsync(SpreadsheetsResource.ValuesResource valuesResource) {
		//	var valueRange = new ValueRange {
		//		Values = new List<IList<object>> {
		//			new List<object> { DateTime.Parse(DateTime.Now.ToString()).ToString("dd/MM/yyyy"),
		//				"Josiel", "E-mail", "Edgar", "SAC (Empreendimentos/DICON)",
		//				DateTime.Parse(DateTime.Now.ToString()).ToString("dd/MM/yyyy"),
		//				DateTime.Parse(DateTime.Now.ToString()).ToString("dd/MM/yyyy"),
		//				"FIX076-2020", "Foi analisado", 100
		//			}
		//		}
		//	};

		//	var update = valuesResource.Update(valueRange, SpreadsheetId, EscreverRange);
		//	update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

		//	var response = await update.ExecuteAsync();
		//	Console.WriteLine($"Updated rows: {response.UpdatedRows}");
		//}

	}
}
