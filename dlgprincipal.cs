using System.ComponentModel.Design.Serialization;
using System.Data;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace PE22B_SGEE
{
    //--------------------------------------------------------------------------------
    //Clase del dialogo principal del programa
    //EESG; 05/09/2022
    //--------------------------------------------------------------------------------
    public partial class dlgprincipal : Form
    {
        //--------------------------------------------------------------------------------
        // Atributos
        //--------------------------------------------------------------------------------
        Color ColorSeleccionado;
        Color ColorSeleccionado2;
        Color ColorSeleccionado3;
        //--------------------------------------------------------------------------------
        // Constructor
        //--------------------------------------------------------------------------------
        public dlgprincipal()
        {
            InitializeComponent();

            ColorSeleccionado = Color.Red;
            ColorSeleccionado2 = Color.AliceBlue;
            ColorSeleccionado3 = Color.Coral;
            P7BtnColor.BackColor = ColorSeleccionado;
            P7BtnColorC.BackColor = ColorSeleccionado2;
            P7BtnColorE.BackColor = ColorSeleccionado3; 
        }
        //--------------------------------------------------------------------------------
        //Boton de saludo
        //--------------------------------------------------------------------------------
        private void Btnsaludar_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("Hola mundo bienvenido a la clase de programcion estructurada", "sistema administrativo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        //--------------------------------------------------------------------------------
        //Saludo de la práctica 1
        //--------------------------------------------------------------------------------
        private void dlgprincipal_Load(object sender, EventArgs e)
        {
            DateTime Fechahoy;
            Fechahoy = DateTime.Now; 
            LblBienvenido.Text = "Bienvenido hoy es " + Fechahoy.ToLongDateString();
        }
        //--------------------------------------------------------------------------------
        //Llenar datos de prueba
        //--------------------------------------------------------------------------------
        private void BtnP2Llenar_Click(object sender, EventArgs e)
        {
            //Limpieza de datos previos
            P2DgvDatos.Rows.Clear();
            P2DgvDatos.Rows.Add();
            //Agrega renglones de datos de prueba
            for (int i = 1; i <= 26; i++)
            {
                P2DgvDatos.Rows.Add();
                P2DgvDatos.Rows[i - 1].Cells[0].Value = i.ToString(); //Codigo
                P2DgvDatos.Rows[i - 1].Cells[1].Value = "Alumno" + i; //Nombre
                if (i > 9)
                {
                    P2DgvDatos.Rows[i - 1].Cells[2].Value = "03-12-20" + i; //Fecha de nacimiento
                    P2DgvDatos.Rows[i - 1].Cells[3].Value = "abc" + i + "1010"; //RFC
                }
                else
                {
                    P2DgvDatos.Rows[i - 1].Cells[2].Value = "03-12-200" + i; //Fecha de nacimiento
                    P2DgvDatos.Rows[i - 1].Cells[3].Value = "abc0" + i + "1010"; //RFC
                }
            }
        }
        //--------------------------------------------------------------------------------
        //Llena datos de prueba desde un archivo de excel
        //--------------------------------------------------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            //Establecer las variables
            DialogResult Resultado;
            OpenFileDialog OpenFileDialog;
            string NombreArchivo;
            string RutaArchivo;
            //Seleccionar el archivo de excel
            OpenFileDialog = new OpenFileDialog();
            OpenFileDialog.Filter = "Archivos de Excel (*.xlsx|*.xlsx|todos los archivos (*.*)|*.*)";
            OpenFileDialog.FilterIndex = 1;
            //Mostrar el archivo seleccionado
            Resultado = OpenFileDialog.ShowDialog();
            //Combrobacion del archivo
            if (Resultado == DialogResult.OK)
            {
                NombreArchivo = OpenFileDialog.SafeFileName.Substring(0, OpenFileDialog.SafeFileName.IndexOf("."));
                RutaArchivo = OpenFileDialog.FileName.Substring(0, OpenFileDialog.FileName.IndexOf(NombreArchivo));
                LeerArchivoExcel( RutaArchivo, NombreArchivo);
            }
        }
        //--------------------------------------------------------------------------------
        //Algoritmo para leer los datos desde un archivo de Excel
        //--------------------------------------------------------------------------------
        private void LeerArchivoExcel(string RutaArchivo, string NombreArchivo)
        {
            //Establecer las variables
            _Application AppExcel;
            _Workbook LibroExcel;
            _Worksheet HojaExcel;
            double Codigo;
            string nombre;
            string ApellidoP;
            string ApellidoM;
            DateTime FechaNac;
            int i;
            //Crear una interpolaridad con Excel
            AppExcel = new Microsoft.Office.Interop.Excel.Application();
            LibroExcel = AppExcel.Workbooks.Open(RutaArchivo + "\\" + NombreArchivo);
            //localiza la hoja de Excel
            HojaExcel = LibroExcel.Sheets["Sheet1"];
            //Limpia datos previos
            P2DgvDatos.Rows.Clear();
            //Encuentra los datos de la tabla mientras se encunetren
            i = 2;
            while(((Range) HojaExcel.Cells[i, 1]).Value != null)
            {
                //Lee los valores de la hoja del excel
                Codigo = (double)((Range)HojaExcel.Cells[i, 1]).Value;
                nombre = (string)((Range)HojaExcel.Cells[i, 4]).Value;
                //Verificar que no esten vacias
                if (((Range)HojaExcel.Cells[i, 2]).Value != null)   
                {
                    ApellidoP = (string)((Range)HojaExcel.Cells[i, 2]).Value;
                }
                else
                {
                   ApellidoP = "";
                }
                if (((Range)HojaExcel.Cells[i, 3]).Value != null)
                {
                    ApellidoM = (string)((Range)HojaExcel.Cells[i, 3]).Value;
                }
                else
                {
                    ApellidoM = "";
                }
                if (((Range)HojaExcel.Cells[i, 5]).Value != null)
                {
                    FechaNac = (DateTime)((Range)HojaExcel.Cells[i, 5]).Value;
                }
                else
                {
                    FechaNac = DateTime.Now;
                }
                //Llena un nuevo renglon de la tabla
                P2DgvDatos.Rows.Add();
                P2DgvDatos.Rows[i - 2].Cells[0].Value = Codigo.ToString();
                P2DgvDatos.Rows[i - 2].Cells[1].Value = nombre;
                P2DgvDatos.Rows[i - 2].Cells[2].Value = ApellidoP;
                P2DgvDatos.Rows[i - 2].Cells[3].Value = ApellidoM;
                P2DgvDatos.Rows[i - 2].Cells[4].Value = FechaNac.ToShortDateString();
                P2DgvDatos.Rows[i - 2].Cells[5].Value = GeneraRFC(ApellidoP, ApellidoM, nombre, FechaNac);
                i++;
            }
            //salir de la funcion
            AppExcel.Quit();
        }
        //--------------------------------------------------------------------------------
        // Tomar valores para Generar RFC
        //--------------------------------------------------------------------------------
        private string GeneraRFC(string ApellidoP, string ApellidoM, string nombre, DateTime FechaNac)
        {
            //Establecer variables
            string RFC = " ";
            string Letra = " ";
            //Obtiene la primera letra del apellido paterno
            RFC = ApellidoP.Substring(0, 1);
            //Obtiene la primera vocal del apellido paterno
            for (int i = 1; i < ApellidoP.Length; i++)
            {
                Letra = ApellidoP.Substring(i, 1);
                if(Letra == "A" || Letra == "E" || Letra == "I" || Letra == "O" || Letra == "U")
                {
                    RFC += Letra;
                    break;
                }
            }
            //Obtiene la primera letra del apellido materno
            if (ApellidoM == "")
            {
                RFC = RFC + "X";
            }
            else
            {
                RFC += ApellidoM.Substring(0, 1);
            }
            //Obtiene la primera letra del nombre
            RFC += nombre.Substring(0, 1);
            //obtiene el año
            RFC +=  FechaNac.ToString("yy");
            //Obtiene el Mes
            RFC +=  FechaNac.ToString("MM");
            //Obtiene el dia
            RFC +=  FechaNac.ToString("dd");
            return RFC;
        }
        //--------------------------------------------------------------------------------
        // Algoritmo para evaluar proyeto de inversion (Vpn)
        //--------------------------------------------------------------------------------
        private void P3BtnCalcular_Click(object sender, EventArgs e)
        {
            //Establecer variables
            int numflujos;
            double Sumatoria;
            double Tasa;
            double Inversion;
            double Resultado;
            double Auxiliar;
            //Comprobar que no esten vacios
            if (P3TxbInverison.Text == " ")
            {
                MessageBox.Show("Capture la inversion inicial");
                P3TxbInverison.Focus();
                return;
            }
            //Combrobar que sea un numero
            if (!double.TryParse(P3TxbInverison.Text, out Auxiliar))
            {
                MessageBox.Show("La inversion inicial debe ser numerica");
                P3TxbInverison.Focus();
                return;
            }
            //Comprobar que no esten vacios
            if (P3TxbTaza.Text== " ")
            {
                MessageBox.Show("Capture la tasa minima esperada");
                P3TxbTaza.Focus();
                return;
            }
            //Combrobar que sea un numero
            if (!double.TryParse(P3TxbTaza.Text, out Auxiliar))
            {
                MessageBox.Show("La Tasa debe ser numerica");
                P3TxbTaza.Focus();
                return;
            }
            //Aplicar las operaciones matematicas
            Tasa = (Double.Parse(P3TxbTaza.Text)/100);
            Inversion = (Double.Parse(P3TxbInverison.Text));
            //Establecer variables
            numflujos = P3DgvFlujosActivos.Rows.Count - 1;
            Sumatoria = 0;
            //Aplicacion de las formulas matematicas
            for (int i = 0; i <numflujos; i++)
            {
                //Variables
                double VF;
                double VP;
                //Mostrar VF en la tabla
                VF = double.Parse( P3DgvFlujosActivos.Rows[i].Cells[1].Value.ToString());
                //Aplicacion matematica
                VP = VF / Math.Pow((1 + Tasa), i+1);
                //Guardar los valores
                Sumatoria = Sumatoria + VP;
            }
            //Guardar los valores
            Resultado = -Inversion + Sumatoria;
            //Formatear el resultado
            P3TxbResultado.Text = string.Format("{0:n}", Resultado);
            //Dar colores a las cajas de texto
            if (Resultado > 0)
            {
                P3TxbResultado.BackColor = Color.LimeGreen;
            }else if(Resultado < 0)
            {
                P3TxbResultado.BackColor = Color.Red;
            }
            else
            {
                P3TxbResultado.BackColor = Color.Yellow;
            }
            //Mostrar la sumatoria
            MessageBox.Show("Sumatoria es =" + Sumatoria);
        }
        private void P3btnLimpiar_Click(object sender, EventArgs e)
        {
            //Limpieza de los valores
            P3TxbAños.Text = "";
            P3TxbInverison.Text = " ";
            P3TxbResultado.Text = "";
            P3TxbTaza.Text = " ";
            P3DgvFlujosActivos.Rows.Clear();
            P3TxbResultado.BackColor = Color.White;
        }
        //--------------------------------------------------------------------------------
        // Boton de ejecucion de Practica 4
        //--------------------------------------------------------------------------------
        private void P4BtnPrueba_Click(object sender, EventArgs e)
        {
            //Llamar a la clase Dlgpractica 4
            DlgPractica4 dlgPractica4;
            dlgPractica4 = new DlgPractica4();
            dlgPractica4.Show();
        }
        //--------------------------------------------------------------------------------
        // boton de llenado de la matriz de la practica 5
        //--------------------------------------------------------------------------------
        private void P5BtnLlenar_Click(object sender, EventArgs e)
        {
            //Variables
            int Renglones;
            int Columnas;
            //Limpiar el contenido de las tablas
            P5DgvTabla.Rows.Clear();
            P5DgvTabla.Columns.Clear();
            //Dar valores a las variables
            Renglones = int.Parse(P5TbxRenglones.Text);
            Columnas = int.Parse(P5TbxColumnas.Text);
            //Agregar las columnas y los renglones con nombre y numero
            for (int i = 0; i < Columnas; i++)
            {
                P5DgvTabla.Columns.Add("Col", "Columna " + (i + 1));
            }
            for (int i = 0; i < Renglones; i++)
            {
                P5DgvTabla.Rows.Add();
            }
        }
        //--------------------------------------------------------------------------------
        // boton de calculo problema de collatz de la practica 5
        //--------------------------------------------------------------------------------
        private void P5BtnCalcular_Click(object sender, EventArgs e)
        {
            //Variables
            long N;
            //Limpiar contenido de las tablas
            P5DgvTabla.Rows.Clear();
            P5DgvTabla.Columns.Clear();
            //Agregar las columnas y los renglones con nombre y numero
            P5DgvTabla.Columns.Add("ColN ", "N");
            N = long.Parse(P5TbxNumeroN.Text);
            //Llamar a la funcion Collatz
            Collatz(N);
            //Mostrar los pasos en el label
            P5LblPasos.Text = (P5DgvTabla.Rows.Count - 1) + " Pasos";
        }
        // Proceso para resolver la secunecia de Collatz de forma recursiva
        private void Collatz(long N)
        {
            //Variables
            long resultado;
            //Agregar N a las tablas
            P5DgvTabla.Rows.Add(N.ToString());
            //Operaciones matematicas de la secuencia de collatz
            if(N != 1)
            {
                if (N % 2 ==0)
                {
                    resultado = N / 2;
                }
                else
                {
                    resultado = (N * 3) + 1;
                }
                Collatz(resultado);
            }
        }
        //--------------------------------------------------------------------------------
        // Practica 6
        // Boton llenar
        //--------------------------------------------------------------------------------
        private void P6BtnLlenar_Click(object sender, EventArgs e)
        {
            //Variables
            int Renglones;
            int Columnas;
            //Limpiar los datos de la tabla
            P6DgvTabla.Rows.Clear();
            P6DgvTabla.Columns.Clear();
            //Dar valores a las variables
            Renglones = int.Parse(P6TbxRenglones.Text);
            Columnas = int.Parse(P6TbxColumnas.Text);
            //Agregar las columnas y los renglones con nombre y numero
            for (int i = 0; i < Columnas; i++)
            {
                P6DgvTabla.Columns.Add("Col", "Columna " + (i + 1));
            }
            for (int i = 0; i < Renglones; i++)
            {
                P6DgvTabla.Rows.Add();
            }
        }
        //--------------------------------------------------------------------------------
        // Boton de calcular Practica 6
        //--------------------------------------------------------------------------------
        private void P6BtnCalcular_Click(object sender, EventArgs e)
        {
            //Variables
            int Renglones;
            int Columnas;
            int R;
            int C;
            Random Aleatorio;
            //Iniciar la variable random
            Aleatorio = new Random();
            //Recupera las dimensiones
            Renglones = int.Parse(P6TbxRenglones.Text);
            Columnas = int.Parse(P6TbxColumnas.Text);
            //llenar con numeros aleatorios
            R = 0;
            while (R <= Renglones - 1)
            {
                C = 0;
                while (C <= Columnas - 1)
                {
                    P6DgvTabla.Rows[R].Cells[C].Value = Aleatorio.Next(100);
                    C++;
                }
                R++;
            }
            R = 0;
            //colorear el triangulo
            while (R <= Renglones - 1)
            {
                C = 0;
                while (C <= Columnas - 1)
                {
                    int auxiliarder ;
                    if (P6ChbInvertir.Checked)
                    {
                        auxiliarder = Renglones - R - 1;
                    }
                    else
                    {
                        auxiliarder = R;
                    }
                    P6DgvTabla.Rows[auxiliarder].Cells[C].Style.BackColor = Color.Yellow;
                    for (int i = 0; i < R; i++)
                    {
                        P6DgvTabla.Rows[auxiliarder].Cells[i].Style.BackColor = Color.White;
                        P6DgvTabla.Rows[auxiliarder].Cells[Columnas - i-1].Style.BackColor = Color.White;
                    }
                    C++;
                }
                R++;
            }
            //Sumatoria de las celdas amarillas
            R = 0;
            long Sumatoria = 0;
            while (R <= Renglones - 1)
            {
                C = 0;
                while (C <= Columnas - 1)
                {
                    if (P6DgvTabla.Rows[R].Cells[C].Style.BackColor == Color.Yellow)
                    {
                        //Capturar el valor de la celda que cumpla la condicion
                        Sumatoria = Sumatoria + long.Parse(P6DgvTabla.Rows[R].Cells[C].Value.ToString());
                    }
                    C++;
                }
                R++;
            }
            //Mostrar mensaje de la sumatoria
            P6TbxSumatoria.Text = Sumatoria.ToString();
        }
        //--------------------------------------------------------------------------------
        // Practica 7
        // Parte 1
        //--------------------------------------------------------------------------------
        private void P7BtnDibujar_Click(object sender, EventArgs e)
        {
            //Variables
            Pen Pluma;
            PointF Punto1;
            PointF Punto2;
            PointF Punto3;
            PointF Punto4;
            Graphics Lienzo;
            float Xinicial;
            float Yinicial;
            float AnchoPluma;
            float Lado;
            //Prepara datos
            Xinicial = float.Parse(P7TbxXinicial.Text);
            Yinicial = float.Parse(P7TbxYinicial.Text);
            AnchoPluma = float.Parse(P7NudPluma.Value.ToString());
            Lado = float.Parse(P7NudLado.Text);
            //Prepara la pluma
            Pluma = new Pen(ColorSeleccionado, AnchoPluma);
            //Calcula Vectores
            Punto1 = new PointF(Xinicial, Yinicial);
            Punto2 = new PointF(Xinicial, Lado + Yinicial);
            Punto3 = new PointF(Lado + Xinicial, Lado + Yinicial);
            Punto4 = new PointF(Lado + Xinicial, Yinicial);
            //Crea Arreglo
            PointF[] VectoresPoligonos =
            {
                Punto1,
                Punto2,
                Punto3,
                Punto4
            };
            //Dibuja el Algoritmo
            Lienzo = P7PnlLienzo.CreateGraphics();
            Lienzo.DrawPolygon(Pluma, VectoresPoligonos);
        }
        //Boton de Color de la pluma de la practica 7
        private void P7BtnColor_Click(object sender, EventArgs e)
        {
            //Inicializar la varible del color
            ColorDialog DlgSelectorColor = new ColorDialog();
            //Ventana de selecion de color
            DlgSelectorColor.AllowFullOpen = true;
            DlgSelectorColor.ShowHelp = true;
            DlgSelectorColor.Color = ColorSeleccionado;
            if (DlgSelectorColor.ShowDialog() == DialogResult.OK)
            {
                //Cambio de color de letra del boton de color
                ColorSeleccionado = DlgSelectorColor.Color;
                P7BtnColor.BackColor = ColorSeleccionado;
                if (ColorSeleccionado.R < 60 && ColorSeleccionado.G < 60 && ColorSeleccionado.B < 60 )
                {
                    P7BtnColor.ForeColor = Color.White;
                }
                else
                {
                    P7BtnColor.ForeColor = Color.Black;
                }
            }
        }
        //Boton de Color de los cuadrados de fibonacci de la practica 7
        private void button1_Click_1(object sender, EventArgs e)
        {
            //Inicializar la varible del color
            ColorDialog DlgSelectorColor = new ColorDialog();
            //Ventana de selecion de color
            DlgSelectorColor.AllowFullOpen = true;
            DlgSelectorColor.ShowHelp = true;
            DlgSelectorColor.Color = ColorSeleccionado2;
            if (DlgSelectorColor.ShowDialog() == DialogResult.OK)
            {
                //Cambio de color de letra del boton de color
                ColorSeleccionado2 = DlgSelectorColor.Color;
                P7BtnColorC.BackColor = ColorSeleccionado2;
            }
        }
        //Boton de Color de la espiral de fibonacci de la practica 7
        private void button2_Click(object sender, EventArgs e)
        {
            //Inicializar la varible del color
            ColorDialog DlgSelectorColor = new ColorDialog();
            //Ventana de selecion de color
            DlgSelectorColor.AllowFullOpen = true;
            DlgSelectorColor.ShowHelp = true;
            DlgSelectorColor.Color = ColorSeleccionado3;
            if (DlgSelectorColor.ShowDialog() == DialogResult.OK)
            {
                //Cambio de color de letra del boton de color
                ColorSeleccionado3 = DlgSelectorColor.Color;
                P7BtnColorE.BackColor = ColorSeleccionado3;
            }
        }
        //--------------------------------------------------------------------------------
        // Practica 7
        // Parte 2
        //--------------------------------------------------------------------------------
        private void P7BtnPractica2_Click(object sender, EventArgs e)
        {
            //Variables y Preparacion de datos
            float Xinicial = P7PnlLienzo.Width / 2, Xfinal;
            float Yinicial = P7PnlLienzo.Height / 2, Yfinal;
            float Lado, LadoFinal;
            float Incremento;
            float NumEspirales;
            float GrosorPluma = ((float)P7NudPluma.Value);
            Pen Pluma = new Pen(ColorSeleccionado, GrosorPluma);
            Graphics Lienzo = P7PnlLienzo.CreateGraphics();
            NumEspirales = float.Parse(P7TbxXinicial.Text);
            Incremento = float.Parse(P7TbxYinicial.Text);
            Lado = float.Parse(P7NudLado.Text);

            for (int i = 0; i < NumEspirales; i++)
            {
                DibujaEspiral(Xinicial, Yinicial, Lado, Incremento, Lienzo, Pluma, out Xfinal, out Yfinal, out LadoFinal);

                Xinicial = Xfinal;
                Yinicial = Yfinal;
                Lado = LadoFinal;
            }
        }
        //Funcion para dibujar espiral de la practica 7
        void DibujaEspiral(float Xinicial, float Yinicial, float Lado, float Incremento, Graphics Lienzo, Pen Pluma, out float Xfinal, out float Yfinal, out float LadoFinal)
        {
            Lado += Incremento;
            PointF Punto1 = new PointF(Xinicial, Yinicial);
            PointF Punto2 = new PointF(Xinicial + Lado, Yinicial);
            PointF Punto3 = new PointF(Xinicial + Lado, Yinicial + Lado);
            PointF Punto4 = new PointF(Xinicial + Incremento, Yinicial + Lado);
            PointF Punto5 = new PointF(Xinicial + Incremento, Yinicial + Incremento);

            PointF[] Espiral =
            {
                Punto1,
                Punto2,
                Punto3,
                Punto4,
                Punto5
            };

            Lienzo.DrawLines(Pluma, Espiral);

            Xfinal = Xinicial - Incremento;
            Yfinal = Yinicial - Incremento;
            LadoFinal = Lado + Incremento;
        }
        //Boton de lanzamiento de la practica 8
        private void P8BtnInicio_Click(object sender, EventArgs e)
        {
            DlgPractica4 dlgPractica4;
            dlgPractica4 = new DlgPractica4();
            dlgPractica4.Show();
        }
        //--------------------------------------------------------------------------------
        // Practica 7
        // Parte 3
        //--------------------------------------------------------------------------------
        private void P7BtnPractica3_Click(object sender, EventArgs e)
        {
            //Variables
            Pen Pluma;
            Pen Pluma3;
            float AnchoPluma;
            float Xinicial;
            float Yfinal;
            float NuevoLado;
            float NuevoLado2;
            float NuevoLado3;
            float NuevoLado4;
            float NuevoLado5;
            float NuevoLado6;
            float nuevoAncho;
            float nuevoAlto;
            float Ancho;
            float RelacionDeOro;
            Graphics Lienzo;
            PointF Punto1, Punto2, Punto3, Punto4, Punto5;
            PointF Punto6, Punto7, Punto8, Punto9, Punto10;
            PointF Punto11, Punto12, Punto13, Punto14, Punto15;
            PointF Punto16, Punto17, Punto18, Punto19, Punto20;
            PointF PuntoC1, PuntoC2, PuntoC3, PuntoC4, PuntoC5, PuntoC6;
            PointF PuntoC7, PuntoC8, PuntoC9, PuntoC10, PuntoC11, PuntoC12;
            PointF PuntoC13, PuntoC14, PuntoC15, PuntoC16, PuntoC17, PuntoC18;
            PointF PuntoC19, PuntoC20, PuntoC21, PuntoC22, PuntoC23, PuntoC24;
            PointF PuntoC25, PuntoC26, PuntoC27, PuntoC28, PuntoC29, PuntoC30;
            PointF PuntoC31, PuntoC32, PuntoC33, PuntoC34, PuntoC35, PuntoC36;
            //Prepara los datos
            AnchoPluma = float.Parse(P7NudPluma.Value.ToString());
            Yfinal = P7PnlLienzo.Height;
            RelacionDeOro = (float)(1 + Math.Sqrt(5)) / 2;
            Ancho = P7PnlLienzo.Height * RelacionDeOro;
            Xinicial = P7PnlLienzo.Height;
            nuevoAncho = Math.Abs(Xinicial - Ancho);
            nuevoAlto = Math.Abs(Xinicial - nuevoAncho);
            NuevoLado = Math.Abs(nuevoAncho - nuevoAlto);
            NuevoLado2 = Math.Abs(nuevoAlto - NuevoLado);
            NuevoLado3 = Math.Abs(NuevoLado - NuevoLado2);
            NuevoLado4 = Math.Abs(NuevoLado2 - NuevoLado3);
            NuevoLado5 = Math.Abs(NuevoLado3 - NuevoLado4);
            NuevoLado6 = Math.Abs(NuevoLado4 - NuevoLado5);
            //Prepara la pluma
            Pluma = new Pen(ColorSeleccionado, AnchoPluma);
            //Calcula Vectores de las lineas de referencia
            Punto1 = new PointF(0, 0);
            Punto2 = new PointF(Ancho,0);
            Punto3 = new PointF(Ancho, Yfinal);
            Punto4 = new PointF(0, Yfinal);
            Punto5 = new PointF(Xinicial, 0);
            Punto6 = new PointF(Xinicial, Yfinal);
            Punto7 = new PointF(Xinicial, nuevoAncho);
            Punto8 = new PointF(Ancho, nuevoAncho);
            Punto9 = new PointF(Ancho - nuevoAlto, Yfinal);
            Punto10 = new PointF(Ancho - nuevoAlto, nuevoAncho);
            Punto11 = new PointF(Xinicial, Yfinal - NuevoLado);
            Punto12 = new PointF(Ancho - nuevoAlto, Yfinal - NuevoLado);
            Punto13 = new PointF(Xinicial + NuevoLado2, Yfinal - nuevoAlto);
            Punto14 = new PointF(Xinicial + NuevoLado2, Yfinal - NuevoLado);
            Punto15 = new PointF(Xinicial + NuevoLado2, nuevoAncho + NuevoLado3);
            Punto16 = new PointF(Xinicial + NuevoLado, nuevoAncho + NuevoLado3);
            Punto17 = new PointF(Ancho - nuevoAlto - NuevoLado4, nuevoAncho + NuevoLado3);
            Punto18 = new PointF(Ancho - nuevoAlto - NuevoLado4, Yfinal - NuevoLado);
            Punto19 = new PointF(Xinicial + NuevoLado2, Yfinal - NuevoLado - NuevoLado5);
            Punto20 = new PointF(Xinicial + NuevoLado2 + NuevoLado5, Yfinal - NuevoLado - NuevoLado5);
            //Crea los Arreglos de las lineas de referencia
            PointF[] VectoresPoligonos =
            {
                Punto1,
                Punto2,
                Punto3,
                Punto4
            };
            PointF[] VectoresPoligonos1 =
            {
                Punto5,
                Punto6
            };
            PointF[] VectoresPoligonos2 =
            {
                Punto7,
                Punto8
            };
            PointF[] VectoresPoligonos3 =
            {
                Punto9,
                Punto10
            };
            PointF[] VectoresPoligonos4 =
            {
                Punto11,
                Punto12
            };
            PointF[] VectoresPoligonos5 =
            {
                Punto13,
                Punto14
            };
            PointF[] VectoresPoligonos6 =
            {
                Punto15,
                Punto16
            };
            PointF[] VectoresPoligonos7 =
            {
                Punto17,
                Punto18
            };
            PointF[] VectoresPoligonos8 =
            {
                Punto19,
                Punto20
            };
            //Dibuja el Algoritmo de las lineas de referencia
            Lienzo = P7PnlLienzo.CreateGraphics();
            Lienzo.DrawPolygon(Pluma, VectoresPoligonos);
            Lienzo.DrawLines(Pluma, VectoresPoligonos1);
            Lienzo.DrawLines(Pluma, VectoresPoligonos2);
            Lienzo.DrawLines(Pluma, VectoresPoligonos3);
            Lienzo.DrawLines(Pluma, VectoresPoligonos4);
            Lienzo.DrawLines(Pluma, VectoresPoligonos5);
            Lienzo.DrawLines(Pluma, VectoresPoligonos6);
            Lienzo.DrawLines(Pluma, VectoresPoligonos7);
            Lienzo.DrawLines(Pluma, VectoresPoligonos8);
            //Prepara la pluma
            Pluma3 = new Pen(ColorSeleccionado, AnchoPluma);
            //Calcula los vectores de las curvas de la figura aurea
            PuntoC1 = new PointF(0F, Yfinal);
            PuntoC2 = new PointF(0F, 0F);
            PuntoC3 = new PointF(Xinicial, 0F);
            PuntoC4 = new PointF(Xinicial, 0F);
            PuntoC5 = new PointF(Xinicial, 0F);
            PuntoC6 = new PointF(Ancho, 0F);
            PuntoC7 = new PointF(Ancho, nuevoAncho);
            PuntoC8 = new PointF(Ancho, nuevoAncho);
            PuntoC9 = new PointF(Ancho, nuevoAncho);
            PuntoC10 = new PointF(Ancho, Yfinal);
            PuntoC11 = new PointF(Ancho - nuevoAlto, Yfinal);
            PuntoC12 = new PointF(Ancho - nuevoAlto, Yfinal);
            PuntoC13 = new PointF(Xinicial + NuevoLado, Yfinal);
            PuntoC14 = new PointF(Xinicial, Yfinal);
            PuntoC15 = new PointF(Xinicial, Yfinal - NuevoLado);
            PuntoC16 = new PointF(Xinicial, Yfinal - NuevoLado);
            PuntoC17 = new PointF(Xinicial, Yfinal - NuevoLado);
            PuntoC18 = new PointF(Xinicial, nuevoAncho);
            PuntoC19 = new PointF(Xinicial + NuevoLado2, nuevoAncho);
            PuntoC20 = new PointF(Xinicial + NuevoLado2, nuevoAncho);
            PuntoC21 = new PointF(Xinicial + NuevoLado2, nuevoAncho);
            PuntoC22 = new PointF(Xinicial + NuevoLado2 + NuevoLado3, nuevoAncho);
            PuntoC23 = new PointF(Xinicial + NuevoLado2 + NuevoLado3, nuevoAncho + NuevoLado3);
            PuntoC24 = new PointF(Xinicial + NuevoLado2 + NuevoLado3, nuevoAncho + NuevoLado3);
            PuntoC25 = new PointF(Xinicial + NuevoLado2 + NuevoLado3, nuevoAncho + NuevoLado3);
            PuntoC26 = new PointF(Xinicial + NuevoLado2 + NuevoLado3, nuevoAncho + NuevoLado3 + NuevoLado4);
            PuntoC27 = new PointF(Xinicial + NuevoLado2 + NuevoLado5, Yfinal - NuevoLado);
            PuntoC28 = new PointF(Xinicial + NuevoLado2 + NuevoLado5, Yfinal - NuevoLado);
            PuntoC29 = new PointF(Xinicial + NuevoLado2 + NuevoLado5, Yfinal - NuevoLado);
            PuntoC30 = new PointF(Xinicial + NuevoLado2, Yfinal - NuevoLado);
            PuntoC31 = new PointF(Xinicial + NuevoLado2, Yfinal - NuevoLado - NuevoLado5);
            PuntoC32 = new PointF(Xinicial + NuevoLado2, Yfinal - NuevoLado - NuevoLado5);
            PuntoC33 = new PointF(Xinicial + NuevoLado2, Yfinal - NuevoLado - NuevoLado5);
            PuntoC34 = new PointF(Xinicial + NuevoLado2, Yfinal - NuevoLado - NuevoLado5 - NuevoLado6);
            PuntoC35 = new PointF(Xinicial + NuevoLado2 + NuevoLado5, Yfinal - NuevoLado - NuevoLado5 - NuevoLado6);
            PuntoC36 = new PointF(Xinicial + NuevoLado2 + NuevoLado5, Yfinal - NuevoLado - NuevoLado5 - NuevoLado6);
            //Crea los arreglos de las curvas aureas
            PointF[] Curva1 =
            {
                PuntoC1,
                PuntoC2,
                PuntoC3,
                PuntoC4
            };
            PointF[] Curva2 =
            {
                PuntoC5,
                PuntoC6,
                PuntoC7,
                PuntoC8
            };
            PointF[] Curva3 =
            {
                PuntoC9,
                PuntoC10,
                PuntoC11,
                PuntoC12
            };
            PointF[] Curva4 =
            {
                PuntoC13,
                PuntoC14,
                PuntoC15,
                PuntoC16

            };
            PointF[] Curva5 =
            {
                PuntoC17,
                PuntoC18,
                PuntoC19,
                PuntoC20

            };
            PointF[] Curva6 =
            {
                PuntoC21,
                PuntoC22,
                PuntoC23,
                PuntoC24

            };
            PointF[] Curva7 =
            {
                PuntoC25,
                PuntoC26,
                PuntoC27,
                PuntoC28

            };
            PointF[] Curva8 =
            {
                PuntoC29,
                PuntoC30,
                PuntoC31,
                PuntoC32

            };
            PointF[] Curva9 =
            {
                PuntoC33,
                PuntoC34,
                PuntoC35,
                PuntoC36

            };
            //Dibuja la curva aurea
            Lienzo = P7PnlLienzo.CreateGraphics();
            Lienzo.DrawBeziers(Pluma3, Curva1);
            Lienzo.DrawBeziers(Pluma3, Curva2);
            Lienzo.DrawBeziers(Pluma3, Curva3);
            Lienzo.DrawBeziers(Pluma3, Curva4);
            Lienzo.DrawBeziers(Pluma3, Curva5);
            Lienzo.DrawBeziers(Pluma3, Curva6);
            Lienzo.DrawBeziers(Pluma3, Curva7);
            Lienzo.DrawBeziers(Pluma3, Curva8);
            Lienzo.DrawBeziers(Pluma3, Curva9);
        }

        //--------------------------------------------------------------------------------
        // Examen (Espiral de Fibonacci)
        //--------------------------------------------------------------------------------
        private void P7btnExamen_Click(object sender, EventArgs e)
        {
            //Limpia el panel del lienzo
            P7PnlLienzo.Refresh();
            //Estlablecer Variables
            Pen Pluma;
            Pen Pluma2;
            Graphics Lienzo;
            float Xcentral;
            float Ycentral;
            float Xfinal;
            float Yfinal;
            float RelacionDeOro;
            float AnchoPluma;
            float AnchoPluma2;
            float Zoom;
            float LadoFinal;
            float X;
            float Y;
            int Vueltas;
            float GrosorCuadros;
            float GrosorEspiral;
            //Establecer los Valores
            AnchoPluma = float.Parse(P7NudCuadros.Value.ToString());
            AnchoPluma2 = float.Parse(P7NudEspiral.Value.ToString());
            Zoom = float.Parse(P7TrbZoom.Value.ToString());
            Zoom = Zoom * 4;
            Vueltas = int.Parse(P7TbxVueltas.Text.ToString());
            RelacionDeOro = (float)(1 + Math.Sqrt(5)) / 2;
            Xcentral = P7PnlLienzo.Width / 2;
            Ycentral = P7PnlLienzo.Height / 2;
            X = float.Parse(P7TbxX.Text.ToString());
            Y = float.Parse(P7TbxY.Text.ToString());
            GrosorCuadros = ((float)P7NudCuadros.Value);
            GrosorEspiral = ((float)P7NudEspiral.Value);
            //Verificar si el usuario quiere mover la espiral
            if (X != 0)
            {
                Xcentral = P7PnlLienzo.Width / 2 + X;
            }
            if (Y != 0)
            {
                Ycentral = P7PnlLienzo.Height / 2 + Y;
            }
            //Preparar Pluma y Lienzo
            Lienzo = P7PnlLienzo.CreateGraphics();
            Pluma = new Pen(ColorSeleccionado2, GrosorCuadros);
            Pluma2 = new Pen(ColorSeleccionado3, GrosorEspiral);
            //Invertir la espiral
            if (P7CbxInvertir.Checked)
            {
               Zoom = -Zoom;
            }
            //Prevenir que las vueltas no superen las 10 (Por limites del lienzo)
            if (Vueltas > 10)
            {
                MessageBox.Show("No Puedes poner mas de 10 vueltas");
                return;
            }
            //Ciclar las vueltas de la espiral segun lo deseado
            for (int i = 0; i < Vueltas; i++)
            {
                //Mandar a llamar la funcion de la espiral
                DibujarEspiralDeFibonacci(Xcentral, Ycentral, RelacionDeOro, AnchoPluma, AnchoPluma2, Zoom, Lienzo, Pluma, Pluma2 ,out Xfinal, out Yfinal, out LadoFinal);
                //Establecer los nuevos datos para las siguientes vueltas
                Ycentral = Yfinal;
                Xcentral = Xfinal;
                Zoom = LadoFinal;
            }
        }
        //Funcion de las espirales de la pactica Examen
        void DibujarEspiralDeFibonacci(float Xcentral, float Ycentral, float RelacionDeOro, float AnchoPluma, float AnchoPluma2, float Zoom, Graphics Lienzo, Pen Pluma, Pen Pluma2, out float Xfinal, out float Yfinal, out float LadoFinal)
        {
            //Crear 1er Cuadro y curva
            PointF Punto1 = new PointF(Xcentral,Ycentral);
            PointF Punto2 = new PointF(Xcentral+Zoom,Ycentral);
            PointF Punto3 = new PointF(Xcentral+Zoom,Ycentral+Zoom);
            PointF Punto4 = new PointF(Xcentral,Ycentral+Zoom);
            PointF[] Cuadro1 =
            {
                Punto1, 
                Punto2, 
                Punto3, 
                Punto4, 
            };
            PointF[] Curva1 =
            {
                Punto1,
                Punto4,
                Punto3,
                Punto3,
            };
            //Dibujar 1er Cuadro y curva
            Lienzo = P7PnlLienzo.CreateGraphics();
            Lienzo.DrawPolygon(Pluma, Cuadro1);
            Lienzo.DrawBeziers(Pluma2, Curva1);
            //Crear 2do Cuadro y curva 
            Punto1 = new PointF(Xcentral+Zoom, Ycentral + Zoom);
            Punto2 = new PointF((Xcentral+Zoom)+(Zoom*RelacionDeOro), Ycentral + Zoom);
            Punto3 = new PointF((Xcentral + Zoom) + (Zoom * RelacionDeOro), (Ycentral + Zoom) - (Zoom * RelacionDeOro));
            Punto4 = new PointF(Xcentral + Zoom, (Ycentral + Zoom) - (Zoom * RelacionDeOro));
            PointF[] Cuadro2 =
            {
                Punto1,
                Punto2,
                Punto3,
                Punto4,
            };
            PointF[] Curva2 =
            {
                Punto1,
                Punto2,
                Punto3,
                Punto3,
            };
            //Dibujar 2do Cuadro y curva 
            Lienzo = P7PnlLienzo.CreateGraphics();
            Lienzo.DrawPolygon(Pluma, Cuadro2);
            Lienzo.DrawBeziers(Pluma2, Curva2);
            //Crear 3er Cuadro y curva 
            Punto1 = new PointF((Xcentral + Zoom) + (Zoom * RelacionDeOro), (Ycentral + Zoom) - (Zoom * RelacionDeOro));
            Punto2 = new PointF((Xcentral + Zoom) + (Zoom * RelacionDeOro), (Ycentral + Zoom) - (Zoom * RelacionDeOro) - ((Zoom * RelacionDeOro) * RelacionDeOro));
            Punto3 = new PointF(Xcentral, (Ycentral + Zoom) - (Zoom * RelacionDeOro) - ((Zoom * RelacionDeOro) * RelacionDeOro));
            Punto4 = new PointF(Xcentral, (Ycentral + Zoom) - (Zoom * RelacionDeOro));
            PointF[] Cuadro3 =
            {
                Punto1,
                Punto2,
                Punto3,
                Punto4,
            };
            PointF[] Curva3 =
            {
                Punto1,
                Punto2,
                Punto3,
                Punto3,
            };
            //Dibujar 3er Cuadro y curva 
            Lienzo = P7PnlLienzo.CreateGraphics();
            Lienzo.DrawPolygon(Pluma, Cuadro3);
            Lienzo.DrawBeziers(Pluma2, Curva3);
            //Crear 4to Cuadro y curva
            Punto1 = new PointF(Xcentral, (Ycentral + Zoom) - (Zoom * RelacionDeOro) - ((Zoom * RelacionDeOro) * RelacionDeOro));
            Punto2 = new PointF(Xcentral - (((Zoom * RelacionDeOro) * RelacionDeOro) * RelacionDeOro), (Ycentral + Zoom) - (Zoom * RelacionDeOro) - ((Zoom * RelacionDeOro) * RelacionDeOro));
            Punto3 = new PointF(Xcentral - (((Zoom * RelacionDeOro) * RelacionDeOro) * RelacionDeOro), Ycentral + Zoom);
            Punto4 = new PointF(Xcentral, Ycentral + Zoom);
            PointF[] Cuadro4 =
            {
                Punto1,
                Punto2,
                Punto3,
                Punto4,
            };
            PointF[] Curva4 =
            {
                Punto1,
                Punto2,
                Punto3,
                Punto3,
            };
            //Dibujar 4to Cuadro y curva 
            Lienzo = P7PnlLienzo.CreateGraphics();
            Lienzo.DrawPolygon(Pluma, Cuadro4);
            Lienzo.DrawBeziers(Pluma2, Curva4);
            //Guardar los datos para las siguientes vueltas
            LadoFinal = -(Xcentral - (((Zoom * RelacionDeOro) * RelacionDeOro) * RelacionDeOro)) + ((Xcentral + Zoom) + (Zoom * RelacionDeOro));
            Xfinal = Xcentral - (((Zoom * RelacionDeOro) * RelacionDeOro) * RelacionDeOro);
            Yfinal = Ycentral + Zoom;
        }
    }//Termina la clase
}//Termina el namespace