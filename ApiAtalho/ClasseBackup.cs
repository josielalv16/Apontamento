//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Threading.Tasks;
//using Google.Apis.Auth.OAuth2;
//using Google.Apis.Services;
//using Google.Apis.Sheets.v4;
//using Google.Apis.Sheets.v4.Data;

//namespace ApiAtalho {
//	class Program {
//		private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
//		private const string SpreadsheetId = "1942C3Hh_V7f9nNzkQ1L-kbtk1KIPg9YTdj86PA7A24s";
//		private const string GoogleCredentialsFileName = "google-credentials.json";
//		/*
//		 * Sheet1 - tab name in a spreadsheet
//		 * A:B - range of values we want to receive
//		*/
//		//private const string ReadRange = "Pagina1!A:I";
//		private static string EscreverRangeApontamento = "A1:J1";
//		private static string EscreverRangeDiario = "A1:J1";

//		static async Task Main(string[] args) {
//			var serviceValues = GetSheetsService().Spreadsheets.Values;

//			// Planilha de Apontamento de Horas Wagner
//			string lerRangeApontamento = "Pagina1!A:I";
//			string idPlanilhaApontamento = "1942C3Hh_V7f9nNzkQ1L-kbtk1KIPg9YTdj86PA7A24s";

//			await LerPlanilha(serviceValues, lerRangeApontamento, idPlanilhaApontamento, false);

//			// Planilha de Apontamento de Horas Diario - Suporte
//			string lerRangeDiario = "Março/2020!A:D";
//			string idPlanilhaDiaria = "1z2euyFZOnmuizYtm-wV-L_5HMS_SkZyNJ7O0oJXqFz8";

//			await LerPlanilha(serviceValues, lerRangeDiario, idPlanilhaDiaria, true);

//			await EscreverApontamento(serviceValues);

//			//await ReadAsync(serviceValues);
//			await WriteAsync(serviceValues);
//		}

//		private static SheetsService GetSheetsService() {
//			using (var stream = new FileStream(GoogleCredentialsFileName,
//				FileMode.Open, FileAccess.Read)) {
//				var serviceInitializer = new BaseClientService.Initializer {
//					HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
//				};
//				return new SheetsService(serviceInitializer);
//			}
//		}

//		private static async Task LerPlanilha(SpreadsheetsResource.ValuesResource valuesResource,
//			string range, string idPlanilha, bool planilhaDiaria) {
//			var response = await valuesResource.Get(idPlanilha, range).ExecuteAsync();
//			var values = response.Values;

//			if (planilhaDiaria)
//				EscreverRangeDiario = $"A{values.Count + 1}:J{values.Count + 1}";
//			else
//				EscreverRangeApontamento = $"A{values.Count + 1}:D{values.Count + 1}";
//		}

//		//private static async Task ReadAsync(SpreadsheetsResource.ValuesResource valuesResource) {
//		//	var response = await valuesResource.Get(SpreadsheetId, ReadRange).ExecuteAsync();
//		//	var values = response.Values;

//		//	if (values == null || !values.Any()) {
//		//		Console.WriteLine("No data found.");
//		//		return;
//		//	}

//		//	var header = string.Join(" ", values.First().Select(r => r.ToString()));
//		//	Console.WriteLine($"Header: {header}");
//		//	foreach (var row in values.Skip(1)) {
//		//		var res = string.Join(" ", row.Select(r => r.ToString()));
//		//		Console.WriteLine(res);
//		//	}

//		//	EscreverRange = $"A{values.Count + 1}:J{values.Count + 1}";
//		//}

//		private static async Task EscreverApontamento(SpreadsheetsResource.ValuesResource valuesResource) {
//			Console.WriteLine("Data da Solicitação: ");
//			string dataSol = Console.ReadLine();
//			Console.WriteLine("Analista:");
//			string analista = Console.ReadLine();
//			Console.WriteLine("Tipo de Atendimento:");
//			string tipoAtend = Console.ReadLine();
//			Console.WriteLine("Solicitante:");
//			string solicitante = Console.ReadLine();
//			Console.WriteLine("Área:");
//			Console.WriteLine("1 - SAC (Empreendimentos/DICON)\n" +
//							  "2 - Comercial (Milton/Beth/Valkiria)\n" +
//							  "3 - Controladoria (Fabio)");
//			int area = int.Parse(Console.ReadLine());
//			Console.WriteLine("Data Inicio:");
//			string dataInicio = Console.ReadLine();
//			Console.WriteLine("Data Fim:");
//			string dataFim = Console.ReadLine();
//			Console.WriteLine("Assunto:");
//			string assunto = Console.ReadLine();
//			Console.WriteLine("Descrição:");
//			string descricao = Console.ReadLine();
//			Console.WriteLine("Tempo Gasto:");
//			string tempo = Console.ReadLine();

//			var valueApontamento = new ValueRange {
//				Values = new List<IList<object>> {
//					new List<object> { dataSol, analista, tipoAtend, solicitante, area}
//				}
//			}


//		}

//		private static async Task WriteAsync(SpreadsheetsResource.ValuesResource valuesResource) {
//			var valueRange = new ValueRange {
//				Values = new List<IList<object>> {
//					new List<object> { DateTime.Parse(DateTime.Now.ToString()).ToString("dd/MM/yyyy"),
//						"Josiel", "E-mail", "Edgar", "SAC (Empreendimentos/DICON)",
//						DateTime.Parse(DateTime.Now.ToString()).ToString("dd/MM/yyyy"),
//						DateTime.Parse(DateTime.Now.ToString()).ToString("dd/MM/yyyy"),
//						"FIX076-2020", "Foi analisado", 100
//					}
//				}
//			};

//			var update = valuesResource.Update(valueRange, SpreadsheetId, EscreverRange);
//			update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

//			var response = await update.ExecuteAsync();
//			Console.WriteLine($"Updated rows: {response.UpdatedRows}");
//		}

//	}
//}
