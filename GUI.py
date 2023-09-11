import wx
import os
import json
import Grade_Statistics as gs


class MyFrame(wx.Frame):
    def __init__(
        self,
        parent,
        id=wx.ID_ANY,
        title="",
        pos=wx.DefaultPosition,
        size=wx.DefaultSize,
        style=wx.DEFAULT_FRAME_STYLE,
    ):
        super(MyFrame, self).__init__(parent, id, title, pos, size, style)

        # 创建一个面板
        panel = wx.Panel(self)

        # 创建一个垂直布局管理器
        rowNum = 10
        sizer = wx.GridBagSizer(rowNum, 2)

        # 创建一个选择文件
        self.selectTemplate = wx.Button(panel, label="选择模版")
        self.selectTemplate.Bind(wx.EVT_BUTTON, self.on_select_template)
        sizer.Add(
            self.selectTemplate,
            pos=(0, 0),
            flag=wx.EXPAND | wx.LEFT | wx.TOP | wx.RIGHT,
            border=10,
        )
        # 创建一个文本框
        self.templateDir = wx.TextCtrl(
            panel,
            style=wx.TE_MULTILINE | wx.ALIGN_CENTER_HORIZONTAL,
            value=f"也可以将模版文件拖拽到这里\n默认为{config['defaultTemplatePath']}",
            size=(100, 40),
        )
        self.templateDir.SetDropTarget(FileDropTarget(self.templateDir))
        # 若文件被拖拽进入则将gs.template设为文件路径
        self.templateDir.Bind(
            wx.EVT_TEXT, lambda event: setattr(gs, "template", event.GetString())
        )
        sizer.Add(
            self.templateDir,
            pos=(1, 0),
            flag=wx.EXPAND | wx.LEFT | wx.TOP | wx.RIGHT,
            border=10,
        )

        # 创建一个选择文件夹的按钮
        self.selectFolder = wx.Button(panel, label="选择装有需要合并的文件的文件夹")
        self.selectFolder.Bind(wx.EVT_BUTTON, self.on_select_folder)
        sizer.Add(
            self.selectFolder,
            pos=(2, 0),
            flag=wx.EXPAND | wx.LEFT | wx.TOP | wx.RIGHT,
            border=10,
        )
        # 创建一个文本框
        self.folderDir = wx.TextCtrl(
            panel,
            style=wx.TE_MULTILINE | wx.ALIGN_CENTER_HORIZONTAL,
            value=f"也可以将文件夹拖拽到这里\n默认为{config['defaultInputPath']}",
            size=(100, 40),
        )
        self.folderDir.SetDropTarget(FileDropTarget(self.folderDir))
        # 若文件被拖拽进入则将gs.path设为文件路径
        self.folderDir.Bind(
            wx.EVT_TEXT, lambda event: setattr(gs, "path", event.GetString())
        )
        sizer.Add(
            self.folderDir,
            pos=(3, 0),
            flag=wx.EXPAND | wx.LEFT | wx.TOP | wx.RIGHT,
            border=10,
        )

        # 创建一个选择输出目录的按钮
        self.selectOutput = wx.Button(panel, label="选择输出目录")
        self.selectOutput.Bind(wx.EVT_BUTTON, self.on_select_output)
        sizer.Add(
            self.selectOutput,
            pos=(4, 0),
            flag=wx.EXPAND | wx.LEFT | wx.TOP | wx.RIGHT,
            border=10,
        )
        # 创建一个文本框
        self.outputDir = wx.TextCtrl(
            panel,
            style=wx.TE_MULTILINE | wx.ALIGN_CENTER_HORIZONTAL,
            value=f"也可以将要输出到的文件夹拖拽到这里\n默认为{config['defaultOutputPath']}",
            size=(100, 40),
        )
        self.outputDir.SetDropTarget(FileDropTarget(self.outputDir))
        # 若文件被拖拽进入则将gs.outputPath设为文件路径
        self.outputDir.Bind(
            wx.EVT_TEXT, lambda event: setattr(gs, "outputPath", event.GetString())
        )
        sizer.Add(
            self.outputDir,
            pos=(5, 0),
            flag=wx.EXPAND | wx.TOP | wx.LEFT,
            border=10,
        )

        # 创建一个多项选择框用于关键字选择
        self.keyWordTitle = wx.StaticText(panel, label="关键字", style=wx.ALIGN_CENTER)
        sizer.Add(
            self.keyWordTitle,
            pos=(0, 1),
            flag=wx.EXPAND | wx.TOP | wx.RIGHT,
            border=20,
        )
        self.keyWord = wx.ComboBox(
            panel,
            choices=config["keyWords"],
            style=wx.CB_DROPDOWN | wx.TE_MULTILINE,
        )
        # 当用户选择或输入时将gs.keyWord设为用户输入
        self.keyWord.Bind(
            wx.EVT_TEXT, lambda event: setattr(gs, "keyWord", event.GetString())
        )
        sizer.Add(
            self.keyWord,
            pos=(1, 1),
            flag=wx.EXPAND | wx.RIGHT,
            border=10,
        )

        # 创建几个多选框用于选择是否生成最高分等信息 默认为选中
        self.checkBoxMax = wx.CheckBox(panel, label="是否生成最高分")
        self.checkBoxMax.SetValue(config["max"])
        # 当用户选择或输入时将gs.maxRow设为用户输入
        self.checkBoxMax.Bind(
            wx.EVT_CHECKBOX, lambda event: setattr(gs, "maxRow", event.IsChecked())
        )
        sizer.Add(
            self.checkBoxMax,
            pos=(2, 1),
            flag=wx.EXPAND | wx.TOP | wx.RIGHT,
            border=10,
        )

        self.checkBoxMean = wx.CheckBox(panel, label="是否生成平均分")
        self.checkBoxMean.SetValue(config["mean"])
        # 当用户选择或输入时将gs.meanRow设为用户输入
        self.checkBoxMean.Bind(
            wx.EVT_CHECKBOX, lambda event: setattr(gs, "meanRow", event.IsChecked())
        )
        sizer.Add(
            self.checkBoxMean,
            pos=(3, 1),
            flag=wx.EXPAND | wx.TOP | wx.RIGHT,
            border=10,
        )

        # 创建一个按钮用于开始合并
        self.mergeButton = wx.Button(panel, label="开始合并")
        self.mergeButton.Bind(wx.EVT_BUTTON, self.on_merge)
        sizer.Add(
            self.mergeButton,
            pos=(4, 1),
            span=(2, 1),
            flag=wx.EXPAND | wx.TOP | wx.RIGHT,
            border=10,
        )

        # 创建一个进度条
        self.gaugeTitle = wx.StaticText(
            panel, label="合并进度", style=wx.ALIGN_CENTER, size=(-1, 20)
        )
        sizer.Add(
            self.gaugeTitle,
            pos=(6, 0),
            span=(1, 2),
            flag=wx.EXPAND | wx.RIGHT | wx.LEFT,
            border=10,
        )
        self.gauge = wx.Gauge(panel, range=100, size=(-1, 10), style=wx.GA_SMOOTH)
        sizer.Add(
            self.gauge,
            pos=(7, 0),
            span=(1, 2),
            flag=wx.EXPAND | wx.RIGHT | wx.LEFT,
            border=10,
        )

        # 创建一个运行日志框
        self.logTitle = wx.StaticText(panel, label="运行日志", style=wx.ALIGN_CENTER)
        sizer.Add(
            self.logTitle,
            pos=(8, 0),
            span=(1, 2),
            flag=wx.EXPAND | wx.RIGHT | wx.LEFT,
            border=10,
        )
        self.log = wx.TextCtrl(
            panel,
            style=wx.TE_MULTILINE | wx.ALIGN_CENTER_HORIZONTAL,
            size=(100, 50),
        )
        sizer.Add(
            self.log,
            pos=(9, 0),
            span=(1, 2),
            flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM,
            border=10,
        )

        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_timer, self.timer)

        sizer.AddGrowableCol(0)
        for row in range(rowNum):
            sizer.AddGrowableRow(row)
        panel.SetSizerAndFit(sizer)

    def on_select_template(self, event):
        wildcard = "Text Files (*.xlsx)|*.xlsx"
        file_dialog = wx.FileDialog(self, "选择姓名模版", wildcard=wildcard, style=wx.FD_OPEN)
        if file_dialog.ShowModal() == wx.ID_CANCEL:
            return
        selected_file = file_dialog.GetPath()
        gs.template = selected_file
        self.templateDir.SetValue(f"{selected_file}")
        print("选择的文件:", selected_file)
        file_dialog.Destroy()

    def on_select_folder(self, event):
        folder_dialog = wx.DirDialog(self, "选择文件夹")
        if folder_dialog.ShowModal() == wx.ID_CANCEL:
            return
        selected_folder = folder_dialog.GetPath()
        gs.path = selected_folder
        self.folderDir.SetValue(f"{selected_folder}")
        print("选择的文件夹:", selected_folder)
        folder_dialog.Destroy()

    def on_select_output(self, event):
        folder_dialog = wx.DirDialog(self, "选择输出目录")
        if folder_dialog.ShowModal() == wx.ID_CANCEL:
            return
        selected_folder = folder_dialog.GetPath()
        gs.outputPath = selected_folder
        self.outputDir.SetValue(f"{selected_folder}")
        print("选择的文件夹:", selected_folder)
        folder_dialog.Destroy()

    def on_merge(self, event):
        self.log.Clear()
        self.gauge.SetValue(0)
        self.timer.Start(100)
        if gs.Merge():
            if gs.template == os.path.join(os.getcwd(), "姓名模版.xlsx"):
                config["defaultTemplatePath"] = "程序目录下的「姓名模版.xlsx」"
            else:
                config["defaultTemplatePath"] = gs.template
            if gs.path == os.path.join(os.getcwd(), "成绩"):
                config["defaultInputPath"] = "程序目录下的「成绩」文件夹"
            else:
                config["defaultInputPath"] = gs.path
            if gs.outputPath == os.getcwd():
                config["defaultOutputPath"] = "程序目录"
            else:
                config["defaultOutputPath"] = gs.outputPath
            if not gs.keyWord in config["keyWords"]:
                config["keyWords"] = [gs.keyWord] + config["keyWords"]
            config["max"] = gs.maxRow
            config["mean"] = gs.meanRow
            with open(
                os.path.join(os.getcwd(), "config", "config.json"),
                "w",
                encoding="utf-8",
            ) as file:
                json.dump(config, file, ensure_ascii=False, indent=4)
        else:
            gs.logs.append("程序异常运行,请查看说明文档和日志以获取帮助")

    def on_timer(self, event):
        while len(gs.logs) != 0:
            self.log.AppendText("\n" + gs.logs.pop(0))
        if gs.gauge != self.gauge.GetValue():
            self.gauge.SetValue(gs.gauge)


class FileDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        super(FileDropTarget, self).__init__()
        self.window = window

    def OnDropFiles(self, x, y, filenames):
        for file in filenames:
            # 在文本框中显示文件路径
            self.window.SetValue(f"{file}")


if __name__ == "__main__":
    try:
        with open(
            os.path.join(os.getcwd(), "config", "config.json"), "r", encoding="utf-8"
        ) as f:
            config = json.load(f)
    except FileNotFoundError:
        config = {
            "defaultTemplatePath": "程序目录下的「姓名模版.xlsx」",
            "defaultInputPath": "程序目录下的「成绩」文件夹",
            "defaultOutputPath": "程序目录",
            "keyWords": ["得分", "选择或输入合并用的关键字"],
            "noneList": ["", "未上传", "nan", "NAN", "NaN"],
        }
        try:
            # 创建一个json文件用于存储用户设置
            with open(
                os.path.join(os.getcwd(), "config", "config.json"),
                "w",
                encoding="utf-8",
            ) as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except FileNotFoundError:
            os.mkdir(os.path.join(os.getcwd(), "config"))
            with open(
                os.path.join(os.getcwd(), "config", "config.json"),
                "w",
                encoding="utf-8",
            ) as f:
                json.dump(config, f, ensure_ascii=False, indent=4)

    if config["defaultTemplatePath"] == "程序目录下的「姓名模版.xlsx」":
        gs.template = os.path.join(os.getcwd(), "姓名模版.xlsx")
    else:
        gs.template = config["defaultTemplatePath"]
    if config["defaultInputPath"] == "程序目录下的「成绩」文件夹":
        gs.path = os.path.join(os.getcwd(), "成绩")
    else:
        gs.path = config["defaultInputPath"]
    if config["defaultOutputPath"] == "程序目录":
        gs.outputPath = os.getcwd()
    gs.keyWord = config["keyWords"][0]
    gs.noneList = config["noneList"]
    app = wx.App(False)
    frame = MyFrame(None, title="每月成绩统计", size=(500, 500))
    frame.Show(True)
    app.MainLoop()
