# Libreria Excel

Libreria hecha para exportar documentos básicos para excel (formato xlsx)

## Librerias de terceros 
*   Microsoft.Office.Interop.Excel
## Hecho por
 * [JhonMontoya] 
 * [JuanYarce] 

## Ejemplo de uso
```
using ExcelLibreria;
using System;
using System.Drawing;

namespace PruebaLibreriaExcel
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ExcelLib _excel = new ExcelLib();
            _excel.Ruta = "D:\\Prueba.xlsx";

            _excel.AddDataString("A1","A1","id");
            _excel.AddDataString("B1", "B1", "nombre");
            _excel.AddDataString("C1", "C1", "fecha nac");
            _excel.AddDataString("D1", "D1", "edad");

            _excel.Style(new StyleGeneric() {
                Begin = "A1",
                End = "D1",
                VerticalAlign = true,
                HorizontalAlign = true,
                WrapText = true,
                Bold = true,
                Color = Color.Orange,
                LineStyle = true,
                LineWeight = 3d });

            for (int i = 2; i < 250; i++)
            {
                persona per = new persona{
                    id =i,
                    nombre =string.Format("Nombre_{0}",i),
                    FechaNacimeinto = DateTime.Now,
                    edad = i
                };
                _excel.AddDataInteger("A" + i,"A"+i,per.id);
                _excel.AddDataString("B" + i, "B" + i, per.nombre);
                _excel.AddDataDateTime("C" + i, "C" + i, per.FechaNacimeinto);
                _excel.AddDataDouble("D" + i, "D" + i, per.edad);
            }
            _excel.AddFormat("C:C", "C:C", "dd/MM/yyyy");
            _excel.SaveBook();
            _excel.CloseBook();
            _excel.CloseApp();
        }
    }
}

class persona
{
    public int id { get; set; }
    public string nombre { get; set; }
    public DateTime FechaNacimeinto { get; set; }
    public double edad { get; set; }
}


```
 
 ### Version: 0.0.1

[JhonMontoya]: <https://github.com/jhonvedo>
[JuanYarce]: <https://github.com/JuanEstebanYC>
  