using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB;
using MongoDB.Driver;
using MongoDB.Bson;
using MongoDB.Driver.Linq;

namespace MongoExcelMigration.Modelos
{
    //Versión: 1.0
    //Fecha: 04 de Abril de 2023
    //Autor: ASLOGIC S.A. DE C.V.
    //Desarrollador: Ammi Jatziry Wang Almazán
    //Módulo: Modelo de MongoDB.
    //Descripción: Clase con la conexión a MongoDB.
    //Historial de cambios:
    //04 de Abril de 2023: Cambios en nombres de algunos cambios.
    public static class mdlMongoDB
    {
        public static void SubirDatos(String sDBNombre,BsonDocument oDocumento)
        {
            try
            {
                MongoClient oClient = new MongoClient("mongodb://localhost:27017");

                var oDatabase = oClient.GetDatabase(sDBNombre);

                var coleccion = oDatabase.GetCollection<BsonDocument>("datos");
                coleccion.InsertOne(oDocumento);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
       
    }
}
