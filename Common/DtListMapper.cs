using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace EPPlusDemo.Common
{
    public class DtListMapper
    {
        //DataTable转List<T>
        public static List<T> DataTableToList<T>(System.Data.DataTable dt) where T : class, new()
        {
            if (dt == null) return null;
            List<T> list = new List<T>();
            //遍历DataTable中所有的数据行  
            foreach (DataRow dr in dt.Rows)
            {
                T t = new T();
                PropertyInfo[] propertys = t.GetType().GetProperties();
                foreach (PropertyInfo pro in propertys)
                {

                    //检查DataTable是否包含此列（列名==对象的属性名）    
                    if (dt.Columns.Contains(pro.Name))
                    {
                        try
                        {
                            if (pro.Name == "year")
                            {
                                object value = dr["plan_date"];

                                //value = Convert.ChangeType(value, pro.PropertyType);//强制转换类型
                                //如果非空，则赋给对象的属性  PropertyInfo
                                if (value != DBNull.Value)
                                {
                                    string year = value.ToString().Trim().Split('-')[0];
                                    pro.SetValue(t, year, null);
                                }
                            }
                            else if (pro.Name == "month")
                            {
                                object value = dr["plan_date"];
                                //value = Convert.ChangeType(value, pro.PropertyType);//强制转换类型
                                //如果非空，则赋给对象的属性  PropertyInfo
                                if (value != DBNull.Value)
                                {
                                    string month = Convert.ToInt32(value.ToString().Trim().Split('-')[1]).ToString();
                                    pro.SetValue(t, month, null);
                                }
                            }
                            else
                            {
                                object value = dr[pro.Name];
                                if (value != null)
                                {
                                    //value = Convert.ChangeType(value, pro.PropertyType);//强制转换类型                                                       
                                    //如果非空，则赋给对象的属性  PropertyInfo
                                    if (value != DBNull.Value)
                                    {
                                        if (pro.PropertyType.FullName == "System.String")
                                        {
                                            pro.SetValue(t, Convert.ToString(value), null);
                                        }
                                        else
                                        {
                                            pro.SetValue(t, Convert.ToDecimal(value), null);
                                        }
                                        //if (pro.PropertyType.FullName == "System.Decimal")
                                        //{
                                        //}
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw;
                        }
                    }
                }
                //对象添加到泛型集合中  
                list.Add(t);
            }
            return list;
        }
    }
}
