using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using ConsoleTables;

namespace ImpArqExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            //no aspnet core é obrigatorio registrar o provider
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);


            using (FileStream stream = File.Open($"{Directory.GetCurrentDirectory()}\\excel\\Arquivo01.xlsx", FileMode.Open, FileAccess.Read))
            {


                //usar esse codigo para carregar arquivo xls
                //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

                //usar esse codigo para carregar arquivo xlsx
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {

                    //essa configuração indica se a primeira linha é um header 
                    //se setar true ele desconsidera a leitura da primeira linha
                    //eu deixo esse cara por padrão false, assim eu consigo montar um dropdown 
                    //com o cabeçalho do arquivo para o usuario poder fazer o de/para entre o arquivo e 
                    //os campos do meu sistema.
                    //como exemplo deixando true
                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
                    };

                    //obtem os dados do excel
                    DataSet ds = excelReader.AsDataSet(conf);

                    List<Cliente> clientes = new List<Cliente>();

                    for (int i = 0; i < ds.Tables["Clientes"].Rows.Count; i++)
                    {
                        var row = ds.Tables["Clientes"].Rows[i];

                        // como eu configurei o que a primeira linha é o header do arquivo, consigo fazer a 
                        // busca pelo nome da coluna
                        clientes.Add(new Cliente
                        {
                            Nome = row["Nome"].ToString(),
                            Sexo = row["Sexo"].ToString(),
                            CEP = row["CEP"].ToString(),
                            Endereco = row["Endereco"].ToString(),
                            Numero = int.Parse(row["Numero"].ToString()),
                            Bairro = row["Bairro"].ToString(),
                            Cidade = row["Cidade"].ToString(),
                            Estado = row["Estado"].ToString()
                        });

                    }

                    ConsoleTable.From(clientes).Write();

                    Console.ReadLine();
                    
                }
            }
            

        }
    }
}
