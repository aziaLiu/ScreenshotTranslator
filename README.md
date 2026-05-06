# 截图翻译

Windows WPF 截图翻译工具。按下自定义全局热键后调用系统截图框，框选区域会进入剪贴板，程序随后执行 Windows OCR，并把识别文本翻译为中文显示在界面上。

## 运行

```powershell
cd .\ScreenshotTranslator
dotnet run
```

## 使用

默认热键是 `Ctrl+Shift+T`。点击热键输入框后按下新的组合键，再点击“保存热键”即可修改。也可以点击“立即截图”手动触发。

识别准确度和“识别语言”强相关。截图前先选择原文语言，例如英文、日文、韩文；程序会按该语言创建 Windows OCR 引擎，并在 OCR 前对较小截图做放大预处理。

## 依赖

OCR 使用 Windows 内置 `Windows.Media.Ocr`，系统需要安装对应识别语言的 OCR 语言包。翻译默认调用 MyMemory 在线接口，因此需要网络连接；后续可替换为 DeepL、Microsoft Translator、OpenAI 或本地翻译服务。
