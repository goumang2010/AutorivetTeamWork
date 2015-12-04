using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace AUTORIVET_KAOHE
{
    interface FormInterface
    {
      // int currentrow;
        //刷新界面
      void rf_gridview(dynamic dt);
      void rf_default();
      void rf_filter();
      DataTable get_datatable();

    }
}
