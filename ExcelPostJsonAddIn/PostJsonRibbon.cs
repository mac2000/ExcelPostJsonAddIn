using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using RestSharp;
using System.Windows.Forms;
using System.Net;

namespace ExcelPostJsonAddIn
{
    public partial class PostJsonRibbon
    {
        private void PostJsonRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;
        }

        private void Application_SheetSelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            buttonSubmit.Enabled = isActiveSelectionInsideTable(Target);
        }

        private bool isActiveSelectionInsideTable(Microsoft.Office.Interop.Excel.Range Target)
        {
            return Target.ListObject != null && !String.IsNullOrEmpty(Target.ListObject.Name);
        }

        private void buttonSubmit_Click(object sender, RibbonControlEventArgs e)
        {
            if (!isValidUrlProvided())
            {
                MessageBox.Show("Provide valid URL to post data to and optionally HTTP Auth credentials", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Range activeCell = (e.Control.Context as Window).Application.ActiveCell;

            if (!isActiveSelectionInsideTable(activeCell)) return;

            RestClient client = new RestClient();
            RestRequest request = new RestRequest(editBoxUrl.Text, Method.POST) { RequestFormat = DataFormat.Json };

            if (isHttpAuthProvided())
                client.Authenticator = new HttpBasicAuthenticator(editBoxUser.Text, editBoxPass.Text);

            request.AddBody(GetCollection(activeCell.ListObject));

            var response = client.Execute(request);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                MessageBox.Show(response.Content, "OK");
            }
            else
            {
                MessageBox.Show(response.Content, response.ErrorMessage, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private bool isHttpAuthProvided()
        {
            return !String.IsNullOrEmpty(editBoxUser.Text) && !String.IsNullOrEmpty(editBoxPass.Text);
        }

        private bool isValidUrlProvided()
        {
            Uri uri;
            return Uri.TryCreate(editBoxUrl.Text, UriKind.Absolute, out uri) && uri.Scheme == Uri.UriSchemeHttp;
        }

        private List<Dictionary<string, object>> GetCollection(ListObject listObject)
        {
            var collection = new List<Dictionary<string, object>>();

            object[,] headers = ToZeroBasedArray(listObject.HeaderRowRange.Value2);
            object[,] data = ToZeroBasedArray(listObject.DataBodyRange.Value2);

            for (int row = 0; row < data.GetLength(0); row++)
            {
                var item = new Dictionary<string, object>();

                for (int column = 0; column < data.GetLength(1); column++)
                {
                    item.Add(headers[0, column].ToString(), data[row, column]);
                }

                collection.Add(item);
            }

            return collection;
        }

        private object[,] ToZeroBasedArray(dynamic data)
        {
            if (!(data is Array)) return new object[1, 1] { { data } };

            object[,] collection = new object[data.GetLength(0), data.GetLength(1)];

            for (int row = 1; row <= data.GetLength(0); row++)
            {
                for (int column = 1; column <= data.GetLength(1); column++)
                {
                    collection[row - 1, column - 1] = data[row, column];
                }
            }

            return collection;
        }
    }
}
