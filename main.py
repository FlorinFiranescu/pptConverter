import comtypes.client

def PPTtoPDF(inputFileName, outputFileName):
    #as found on https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
    pptToPdfFormatType = 32
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, pptToPdfFormatType)
    deck.Close()
    powerpoint.Quit()

def DOCtoPDF(inputFileName, outputFileName):
    #as found on https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
    docToPdfFormatType = 17
    word = comtypes.client.CreateObject("Word.Application")
    help(word.Documents)

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = word.Documents.Open(inputFileName)
    deck.SaveAs(outputFileName, docToPdfFormatType)
    deck.Close()
    word.Quit()

pptInFile = "./samplePptx.pptx"
pdf2pptOutFile = "./samplePptxTest"
docInFile = "./sampleWord.docx"
pdf2docOutFile = "./sampleWordTest"
PPTtoPDF(pptInFile, pdf2pptOutFile)
DOCtoPDF(docInFile, pdf2docOutFile)