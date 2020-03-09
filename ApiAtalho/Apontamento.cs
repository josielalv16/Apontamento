using System;
using System.Collections.Generic;
using System.Text;

namespace ApiAtalho {
	public class Apontamento {
		public string data;
		public string pessoa;
		public string assunto;
		public string minutos;

		public Apontamento(string data, string pessoa, string assunto, string minutos) {
			this.data = data;
			this.pessoa = pessoa;
			this.assunto = assunto;
			this.minutos = minutos;
		}
	}
}
