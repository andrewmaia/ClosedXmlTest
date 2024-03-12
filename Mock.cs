using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ClosedXmlTest
{
    public class Mock
    {
        public static List<Pessoa> MockarBaseDadosSaude(){
            var l = new List<Pessoa>
            {
                new("Empresa do Joao", "41.646.207/0001-15", "Masculino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500),
                new("Empresa da Maria", "41.646.207/0001-15", "Feminino", "Identificacao", new DateTime(2000, 1, 1), 24, "Faixa", "Pai", "Situacao", "11111111", "Santos", "SP", "Bradesco", "Plano Bradesco", 500)
            };
            return l;

        }

        public static List<SubEstipulante> MockarSubEstimulantes(){
            var l = new List<SubEstipulante>
            {
                new("Empresa do Joao", "41.646.207/0001-15")
                ,new("Empresa da Maria", "41.646.207/0001-14")
                ,new("Empresa do José", "41.646.207/0001-16")
            };
            return l;

        }        
    }

    public record Pessoa(string Empresa, string CNPJ, string Sexo, string Identificacao, DateTime DataNascimento, int Idade, string FaixaEtaria, string Parentesto, string Situacao, string CID, string Municipio, string UF, string Operadora, string Plano, int ValorAtual);
    public record SubEstipulante(string RazaoSocial, string CNPJ);    
}