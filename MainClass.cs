namespace Assignment_Information_Encoding_4
{
    class MainClass
    {
        static void Main(string[] args)
        {

            //-------------------------------------------Making objecct of WordDocument Class file--------------------------------------------
            WordDocument wordDocument = new WordDocument();
            wordDocument.WritingInWordDocument();


            //-------------------------------------------Making objecct of SpreadSheet Class file--------------------------------------------
            SpreadSheet ss = new SpreadSheet();
                ss.creatingSpreadSheet();

            //-------------------------------------------Making objecct of PPTFile Class file---------------------------------------------------------
            PPTFile pptFile = new PPTFile();
            string filepath = @"C:\\Users\\Public\\Music\\info.pptx";
            pptFile.CreatePresentation(filepath);
        }      
        
    }
} 
    
