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
        public string[][] wordsAndSyn { get; set; }

        public void OnGet()
        {
            Message = "Your application description page.";
            words = new string[0];
            wordsAndSyn = new string[0][];
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
            string text = Request.Form["text1"];

            words = text.Split(' ');

            string[] duplicateWords = GetDuplicatesInStringArray(words);


            List<List<string>> wordList = new List<List<string>>();

            foreach(string word in duplicateWords)
            {
                List<string> oneWord = new List<string>();
                oneWord.Add(word);
                string[] synList = GetSynonyms(word).ToArray();
                synList = RemoveDuplicatesInStringArray(synList);

                foreach(string syn in synList)
                {
                    oneWord.Add(syn);
                }
                wordList.Add(oneWord);
            }

            wordsAndSyn = Convert2DListToArrayOfArrays(wordList);

        }

        public string[] RemoveDuplicatesInStringArray(string[] words)
        {
            List<string> wordList = new List<string>(words);
            for(int i = 0; i < wordList.Count; i++)
            {
                for(int j = 0; j < wordList.Count; j++)
                {
                    if ((i != j) && (wordList[i].Equals(wordList[j])))
                    {
                        wordList.RemoveAt(i);
                    }
                }
            }
            string[] arrayWithoutDuplicates = wordList.ToArray();
            return arrayWithoutDuplicates;
        }

        public string[] GetDuplicatesInStringArray(string[] words)
        {
            List<string> wordList = new List<string>(words);
            List<string> duplicatesList = new List<string>();


            for(int i = 0; i < wordList.Count; i++)
            {
                for(int j = 0; j < wordList.Count; j++)
                {
                    if((i != j) && (wordList[i].Equals(wordList[j])))
                    {
                        duplicatesList.Add(wordList[i]);
                        wordList.RemoveAt(i);
                    }
                }
            }
            string[] duplicatesArray = duplicatesList.ToArray();
            return duplicatesArray;
        }

        public string[][] Convert2DListToArrayOfArrays(List<List<string>> listTwo)
        {
            string[][] arrayOfArrays = new string[listTwo.Count][];
            int index = 0;
            foreach (var x in listTwo)
            {
                arrayOfArrays[index] = x.ToArray();
                index++;
            }
            return arrayOfArrays;
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
