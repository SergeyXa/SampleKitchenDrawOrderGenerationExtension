﻿using KD.SDK;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ord_kdla
{
    public class Order
    {
        Appli _appli = new Appli();
        Scene _scene = new Scene();

        public void WriteLineToCsv(TextWriter textWriter, params string[] values)
        {
            textWriter.WriteLine(
                string.Join(",",
                    values.Select(value => 
                        value.Contains(",") 
                        ? $"\"{value}\"" 
                        : value)));
        }

        public bool GenerateOrder(int unused)
        {
            try
            {
                bool isFactorized = true;

                string orderFilename = _appli.GetCallParamsInfo(AppliEnum.CallParamId.ORDERFILENAME);
                string supplierId = _appli.GetCallParamsInfo(AppliEnum.CallParamId.SUPPLIERID);

                var dataForCsv = new List<(string Reference, string Quantity, string Height, string Width, string Depth, string ModelName)>();

                int objectsNb = _scene.SupplierGetObjectsNb(supplierId, -1, isFactorized);

                var textWriter = new StreamWriter(orderFilename);

                for (int rank = 0; rank < objectsNb; rank++)
                {
                    int objectId = _scene.SupplierGetObjectId(supplierId, -1, isFactorized, rank);

                    string reference = _scene.ObjectGetInfo(objectId, SceneEnum.ObjectInfo.REF);
                    string quantity =_scene.ObjectGetInfo(objectId, SceneEnum.ObjectInfo.QUANTITY);
                    string orderHeight =_scene.ObjectGetInfo(objectId, SceneEnum.ObjectInfo.ORDERDIMZ);
                    string orderWidth =_scene.ObjectGetInfo(objectId, SceneEnum.ObjectInfo.ORDERDIMX);
                    string orderDepth = _scene.ObjectGetInfo(objectId, SceneEnum.ObjectInfo.ORDERDIMY);
                    string modelName = _scene.ObjectGetInfo(objectId, SceneEnum.ObjectInfo.MODELNAME);

                    WriteLineToCsv(textWriter,
                        reference, quantity, orderHeight, orderWidth, orderDepth, modelName);
                }

                textWriter.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

                return false;
            }

            return true;
        }
    }
}