using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace wcfExel
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract]
    public interface IExcel
    {
        [OperationContract]
        byte[] datatableToExcel(string[] Baslik, DataTable dataTable, Tema t, int[] toplam);

        [OperationContract]
        byte[] jsonToExcel(string json, Tema t = Tema.Mavi, int[] toplam = null);
    }
}
