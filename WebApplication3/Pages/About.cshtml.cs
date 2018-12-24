using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace WebApplication3.Pages
{
    public class AboutModel : PageModel
    {
        public string Message { get; set; }
        public string text { get; set; }
        public string[] words { get; set; }
        public string[,] wordsAndSyn { get; set; }
        public int outerIndex { get; set; }
        public int innerIndex { get; set; }
        public List<List<string>> test { get; set; }

        public void OnGet()
        {
            Message = "Your application description page.";
            words = new string[0];
            wordsAndSyn = new string[0,0];
            test = new List<List<string>>();
            outerIndex = 0;
            innerIndex = 0;
            /**
            @{
                var totalMessage = "";
                if (IsPost)
                {
                    var num1 = Request["text1"];
                    var num2 = Request["text2"];
                    var total = num1.AsInt() + num2.AsInt();
                    totalMessage = "Total = " + total;
                }
            }
            **/
        }

        public void OnPost()
        {
            string text1 = Request.Form["text1"];

            text = text1;

            words = text1.Split(' ');

            List<List<string>> testList = new List<List<string>>();

            foreach(string word in words)
            {
                List<string> oneWord = new List<string>();
                oneWord.Add(word);
                foreach(string syn in GetSynonyms(word))
                {
                    oneWord.Add(syn);
                }
                testList.Add(oneWord);
            }
            // convert it to a regular array when I go over so I can deal with getting element [0]
            test = testList;
            /**
            foreach (string word in words)
            {
                innerIndex = 0;
                wordsAndSyn[outerIndex, innerIndex] = word;
                innerIndex += 1;
                foreach (var syn in GetSynonyms(word))
                {
                    wordsAndSyn[outerIndex, innerIndex] = syn;
                    innerIndex += 1;
                }
                outerIndex += 1;
            }
            **/



            //foreach (string word in words)
            //{
            //    text += word + " ";
            //}

            /**
            text = "";
            foreach(var word in words)
            {
                text += ".random.";
                foreach (var value in GetSynonyms(word))
                {
                    text += value + " ";
                }
            }
            **/

        }

        public IEnumerable<string> GetSynonyms(string term)
        {
            var appWord = new Microsoft.Office.Interop.Word.Application();
            object objLanguage = Microsoft.Office.Interop.Word.WdLanguageID.wdEnglishUS;
            Microsoft.Office.Interop.Word.SynonymInfo si = appWord.get_SynonymInfo(term, ref (objLanguage));
            foreach (var meaning in (si.MeaningList as Array))
            {
                yield return meaning.ToString();
            }
            appWord.Quit(); //include this to ensure the related process (winword.exe) is correctly closed. 
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appWord);
            objLanguage = null;
            appWord = null;
        }

    }
}
