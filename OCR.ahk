#NoEnv
SetWorkingDir %A_ScriptDir%

F4::
    pToken := Gdip_Startup()
    pBitmap := Gdip_BitmapFromScreen("0|0|200|200")
    Gdip_SaveBitmapToFile(pBitmap, "temp.png")
    Gdip_DisposeImage(pBitmap)
    Gdip_Shutdown(pToken)

    text := WindowsOCR("temp.png")
    if (text != "")
    {
        MsgBox, 텍스트를 찾았습니다: %text%
    }
    FileDelete, temp.png
return

WindowsOCR(imagePath) {
    static OcrEngine := ComObject("Windows.Media.Ocr.OcrEngine")
    static BitmapDecoder := ComObject("Windows.Graphics.Imaging.BitmapDecoder")
    static SoftwareBitmap := ComObject("Windows.Graphics.Imaging.SoftwareBitmap")

    lang := OcrEngine.AvailableRecognizerLanguages[0]
    ocrEngine := OcrEngine.TryCreateFromLanguage(lang)

    uri := ComObject("System.Uri")
    uri.OriginalString := imagePath

    streamReference := ComObject("Windows.Storage.Streams.RandomAccessStreamReference")
    reference := streamReference.CreateFromUri(uri)

    ReadAsync := reference.OpenReadAsync()
    stream := ReadAsync.GetResults()

    CreateAsync := BitmapDecoder.CreateAsync(stream)
    decoder := CreateAsync.GetResults()

    GetSoftwareBitmapAsync := decoder.GetSoftwareBitmapAsync()
    softwareBitmap := GetSoftwareBitmapAsync.GetResults()

    RecognizeAsync := ocrEngine.RecognizeAsync(softwareBitmap)
    ocrResult := RecognizeAsync.GetResults()

    return ocrResult.Text
}