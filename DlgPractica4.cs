using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace PE22B_SGEE
{
    public partial class DlgPractica4 : Form
    {
        //--------------------------------------------------------------------------------
        // Constructor
        //--------------------------------------------------------------------------------
        public DlgPractica4()
        {
            InitializeComponent();
        }
        //--------------------------------------------------------------------------------
        // Obtiene la coordenada del lugar seleccionado a travez de la API de Google Maps
        //--------------------------------------------------------------------------------
        private async void P4BtnObtenerCoordenadas_Click(object sender, EventArgs e)
        {
            //Valida los datos de entrada
            if ( P4TbxLugar.Text == "")
            {
                return;
            }
            //Obtiene coordenadas
            await GetCoordenadas();
        }
        //--------------------------------------------------------------------------------
        // Obtiene la coordenada de manera asincrona
        //--------------------------------------------------------------------------------
        public async Task GetCoordenadas()
        {
            //Variables
            HttpClient ClienteHttp;
            Uri Direccion;
            HttpResponseMessage RespuestaHttp;
            string ContenidoHttp;
            string Descripcion;
            string Latitud;
            string Longiud;
            string Llave;
            string Lugar;
            string Status;
            XmlDocument DocumentoXml;
            XmlNodeList elemList;
            XmlElement bookElement;
            //Prepara datos de Trabajo
            Llave = "AIzaSyAqPyie1EMOOceRXH7Nk7fSUBoxuhNv9wI";
            Lugar = P4TbxLugar.Text;
            //Consulta la Api de Geolocalizacion de Google Maps
            ClienteHttp = new HttpClient();
            Direccion = new Uri("https://maps.googleapis.com/maps/api/geocode/");
            ClienteHttp.BaseAddress = Direccion;
            //Recibe respuesta de la Api de geolocalizacion de Google Maps
            RespuestaHttp = await ClienteHttp.GetAsync("xml?address=" + Lugar + "&key=" + Llave);
            ContenidoHttp = await RespuestaHttp.Content.ReadAsStringAsync();
            //Procesa la respuesta Xml
            DocumentoXml = new XmlDocument();
            DocumentoXml.LoadXml(ContenidoHttp);
            elemList = DocumentoXml.GetElementsByTagName("status");
            bookElement = (XmlElement)elemList[0];
            Status = bookElement.InnerText;
            elemList = DocumentoXml.GetElementsByTagName("formatted_address");
            bookElement = (XmlElement)elemList[0];
            Descripcion = bookElement.InnerText;
            elemList = DocumentoXml.GetElementsByTagName("location");
            bookElement = (XmlElement)elemList[0];
            Latitud = bookElement["lat"].InnerText;
            Longiud = bookElement["lng"].InnerText;
            //Representa los datos al usuario
            P4TbxLatitud.Text = Latitud;
            P4TbxLongitud.Text = Longiud;
            P4TbxDescripcion.Text = Descripcion;
        }
        //--------------------------------------------------------------------------------
        // Algoritmo para convertir coordenadas a GMS
        //--------------------------------------------------------------------------------
        private void P4BtnConvertir_Click(object sender, EventArgs e)
        {
            //Variables
            double Latitud;
            double Longitud;
            int Grados;
            int Minutos;
            double Segundos;
            //Asigna los valores a las variables
            Latitud = double.Parse(P4TbxLatitud.Text);
            Longitud = double.Parse(P4TbxLongitud.Text);
            //Llama a la funcion para la conversion de coordenadas
            ProcesaConversiondeCoordenadas(Latitud, out Grados, out Minutos, out Segundos, P4TbxLatitudGMS, P4CbxNorte, P4CbxSur );
            ProcesaConversiondeCoordenadas(Longitud, out Grados, out Minutos, out Segundos, P4TbxLongitudGMS, P4CbxEste, P4CbxOeste);
        }
        //--------------------------------------------------------------------------------
        // Algoritmo para convertir coordenadas decimales a GMS
        //--------------------------------------------------------------------------------
        private void ConvertirDecimalesaGMS(double valor, out int Grados, out int minutos, out double segundos)
        {
            //Operaciones matematicas para la conversion
            Grados = (int)valor;
            valor = Math.Abs(valor - Grados);
            minutos = (int)(valor = 60);
            valor = valor - (double)minutos / 60;
            segundos = (double)(valor * 3600);
        }
        //--------------------------------------------------------------------------------
        // Proceso de conversion de coordenadas
        //--------------------------------------------------------------------------------
        private void ProcesaConversiondeCoordenadas (double valor,out int Grados,out int Minutos,out double segundos, TextBox tbxReceptor, CheckBox Positivo , CheckBox Negativo )
        {
            //Regresar los valores
            ConvertirDecimalesaGMS(valor, out Grados, out Minutos, out segundos);
            //Marcardo de los Checkbox segun las coordenadas
            if (Grados > 0)
            {
                Positivo.Checked = true;
                Negativo.Checked = false;
            }
            else
            {
                Positivo.Checked = false;
                Negativo.Checked = true;
            }
            //Establecer formato de la caja de texto
            tbxReceptor.Text = Math.Abs(Grados) + "Grados " + Minutos + "Minutos " + Math.Round(segundos, 4) + "Sengundos ";
        }
        //--------------------------------------------------------------------------------
        // Genera el archivo KML del lugar buscado
        //--------------------------------------------------------------------------------
        private void P8BtnGenerarKML_Click(object sender, EventArgs e)
        {
            //Declaracion de variables
            StreamWriter outputfile;
            //Seleccionar la ruta de creacion del archivo
            //Varia segun la computadora
            //outputfile = new StreamWriter("C:\\Users\\Eduardo\\Desktop\\"+P4TbxLugar.Text+".kml");
            outputfile = new StreamWriter("C:\\Users\\217437526\\Desktop\\"+P4TbxLugar.Text+".kml");
            //Aplicacion de variables en el formato del KML
            string[] lineas = {
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<kml xmlns=\"http://www.opengis.net/kml/2.2\">",
                "  <Placemark>",
                "    <name>",
                "        " + P4TbxLugar.Text,
                "    </name>",
                "    <description>",
                "         " + P4TbxDescripcion.Text,
                "    </description>",
                "    <Point>",
                "\t<extrude>",
                "           1",
                "        </extrude>",
                "\t<altitudeMode>",
                "           relativeToGround",
                "        </altitudeMode>",
                "\t<coordinates>",
                "           " + P4TbxLongitud.Text  + "," + P4TbxLatitud.Text + "," + (TrbAltura.Value * 100),
                "        </coordinates>",
                "    </Point>",
                "  </Placemark>",
                "</kml>" };
            //Escritura del archivo
            foreach (string linea in lineas)
            {
                outputfile.WriteLine(linea);
            }
            outputfile.Close(); 
        }
    }
}
