// Criador: Matheus Yan Monteiro

using System;
using System.Net;
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace download_de_arquivos
{
    class Program
    {
        // classes 
        public class ProcessoSeletivo
        {
            public int CodProjeto { get; set; }
            public string Titulo { get; set; }
            public double Valor { get; set; }
        }

        // funções 
        private static async Task SaveExcelFile(FileInfo content, List<ProcessoSeletivo> popula)
        {
            DeleteIfExists(content);

            using var package = new ExcelPackage(content);

            var worksheet = package.Workbook.Worksheets.Add("Processo2021");

            var range = worksheet.Cells["A1"].LoadFromCollection(popula, true);
            range.AutoFitColumns();

            await package.SaveAsync();

        }
        static List<ProcessoSeletivo> GetSetupData()
        {
            List<ProcessoSeletivo> output = new()
            {
                new() { CodProjeto = 1, Titulo = "Projeto 1", Valor = 100000 },
                new() { CodProjeto = 2, Titulo = "Projeto 2", Valor = 550000 },
                new() { CodProjeto = 3, Titulo = "Projeto 3", Valor = 550000 }
            };

            return output;
        }


        private static void DeleteIfExists(FileInfo content)
        {
            if (content.Exists)
                content.Delete();
        }

        // inicializador do programa 
        static async Task Main(string[] args)
        {
            // bliblioteca que me permite fazer o download da planilha para minha pasta de soluções _src. 
            WebClient DonwloadPath = new WebClient(); //instancia da classe da bliblioteca. 
            DonwloadPath.DownloadFile("https://docs.google.com/spreadsheets/d/1pv_DzWXGRkm4pimTbJomf_8U2UbzIZ_uTg3y31XqURo/edit#gid=0",
                                      @"C:\_src\Processos Seletivos 2021.xlsx"); // chamando o meu metodo de download.

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var content = new FileInfo(@"C:\_src\Processos Seletivos 2021.xlsx");

            var popula = GetSetupData();


            await SaveExcelFile(content, popula);
        }

        
    }
}
