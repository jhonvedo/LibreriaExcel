# ExcelLibrary.Reader

## Librerias de terceros 
*   

## Hecho por
 * [JhonMontoya] 
 * [JuanYarce] 

## Ejemplo de uso
```
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CargueExcel
{
    class Test
    {
        static void Main(string[] args)
        {
            ReaderExcel Cargue = new ReaderExcel();
            Cargue.Path = "D:\\Prueba.xlsx";
            Cargue.SheetSearch(0);
            ActionReturn objeto = Cargue.Load();

            if (objeto.ListData.Count == 0)
            {
                switch (objeto.Accion)
                {
                    case ReaderExcel.ERROR_CONNECTION:
                        Console.WriteLine("error conexion");
                        Console.ReadLine();
                        break;

                    case ReaderExcel.HOJA_NO_ENCOTRADA:
                        Console.WriteLine("hoja no encontrada");
                        Console.ReadLine();
                        break;

                    case ReaderExcel.LIBRO_NO_ENCONTRADO:
                        Console.WriteLine("libro no encontrada");
                        Console.ReadLine();
                        break;
                }
            }
            foreach (var fila in objeto.ListData)
            {
                string texto = "";
                foreach (var column in fila)
                {
                    texto += "-" + column + "-";
                }
                Console.WriteLine(texto);
            }
            Console.ReadLine();
        }
    }
}
```

[JhonMontoya]: <https://github.com/jhonvedo>
[JuanYarce]: <https://github.com/JuanEstebanYC>
  