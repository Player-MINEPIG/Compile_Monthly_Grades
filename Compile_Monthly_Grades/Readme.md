# 诡异的月底成绩合并器v1.0

## 目录

#### 1. 安装

#### 2. 使用说明

#### 3. 运行日志

#### 4. 碎碎念

---

### 1. 安装

- **自动安装** (适用于电脑中未配置python环境的用户)

  提供整合包, 所有依赖全部都装在一个exe文件中, 点开即用

- **手动安装** (适用于电脑中已经配置过python环境, 不想浪费空间的用户)

  0. 创建并进入虚拟环境 (如果需要的话, 如果不清楚可以自行查找虚拟环境的作用)

     ```
     conda create -n mergeExcel
     conda activate mergeExcel
     ```

  1. 下载源码

     - 在github上下载zip, 并解压

     - 或`git clone`

  2. 安装依赖项

     1. 运行终端 (windows中快捷键win+R在弹窗中填入`cmd`)

     2. 在终端中输入`cd yourDirect`将`yourDirect`替换为源码所在文件夹路径并回车 (此时终端中光标前应显示出文件夹路径)

     3. 输入`pip install -r requirements.txt`

  3. 运行`GUI.py`即可

---

### 2. 使用说明

页面如图所示

![image-20230911200336343](/Users/pmp/Library/Application Support/typora-user-images/image-20230911200336343.png)

包含

- **模版**

  **文件后缀名**

  - 模版文件应为**xlsx或xls文件**

  **文件内格式**

  - 形如下图

    ![image-20230911202423431](/Users/pmp/Library/Application Support/typora-user-images/image-20230911202423431.png)

    **必须包含**

    - 一个姓名列
    - 或一个学号列

    可以同时拥有, 若同时拥有则要求在同一行 (不要求在第一行, 但程序会自动去掉在姓名和学号上面的几行), 优先根据学号合并, 防止同名问题 (同名会导致数据重复一次)

    **推荐包含**

    - 放在**第一列**的「**班级最高分**」或「**班级平均分**」, **如果放在其他列同时勾选生成则仍会在第一列创建新的最高分和平均分**

    **可以包含**

    - 其它列, 如第三列已经有数据, 仍可以合并

    **不能包含**

    - 原来有数据后来被删掉的空列, 可能会引发不知名的bug, 虽然已经修复了一次, 但不保证不会再产生
    - 作者还没测试出来, 期待发现~~ (最好别发现)~~

    - 总之按照提供的图片那样总不会错

  **添加方式**

  - 用按钮选择文件
  - 或直接拖入文件

- **需合并的文件**

  **文件后缀名**

  - 因为原本是只用来合并从那个分数系统里面统一下载下来的文件, 所以仅支持**xlsx或xls格式**, 即选择文件夹时只会合并其中的xlsx和xls文件

  **文件内格式**

  - 如下图所示

    ![image-20230911210711344](/Users/pmp/Library/Application Support/typora-user-images/image-20230911210711344.png)

    **必须包含**

    - 一个姓名列
    - 或一个学号列

    可以同时拥有, 优先根据学号合并, 防止同名问题 (同名会导致数据重复一次)

    - 统一的关键字, 如此处从系统下载的文档默认分数栏标题都为**得分**

    **推荐包含**

    - 无

    **可以包含**

    - 其它标题与姓名/学号/关键字 (得分)无关的列

    **不能包含**

    - **多列以姓名/学号/关键字 (得分)**为标题的数据
    - 作者还没测试出来, 期待发现~~ (最好别发现)~~

    - 总之直接用下载下来的原文件就没什么问题应该

  **添加方式**

  - 可以直接拖入文件夹或用按钮选择一个文件夹
  - 可以直接拖入文件, 但是无法用按钮选择文件

- **输出目录**

  **选择方式**

  - 用按钮选择文件夹
  - 拖入文件夹 **(请注意不要拖入文件, 如果拖入文件, 会将结果输出到程序同目录下)**

  **输出形式**

  - 以一个名为**成绩汇总.xlsx**的文件输出到指定文件夹, 不会改变模版或进行合并操作的文件, 若指定目录已存在同名文件则会覆盖同名文件

- 关键字

  **选择方式**

  - 多选框中进行多选
  - 手动输入

  **格式**

  - 同需要合并的部分标题一致, 以上图为例如需要合并分数, 则使用**得分**作为关键字, 若合并班级排名, 则使用**班级名次**作为关键字

- 是否自动生成最高/平均分
  - 若模版第一列中不存在「班级平均分」或「班级最高分」则会在生成时自动在第一列末尾添加
  - 生成的是excel中的公式而非数据

- 开始合并
  - 作用即为开始合并所示

- 合并进度
  - 用于显示合并是否正常完成

- 运行日志
  - 输出运行状态, 详见**运行日志**篇



#### 注意：程序在完整成功运行后会自动保存上一次的设置（包括添加关键字进入选项），保存在程序目录下的config文件夹下的config.json文件中，可以通过手动更改json文件或者将其完整运行一次来更改下一次打开程序的默认状态（若要删除关键字选项则只能手动更改json）

---

### 3. 运行日志

- **文件目录提示** (按运行时输出的顺序排序)
  - 姓名模版路径
  - 成绩文件夹路径
  - 要合并的文件名
  - 要合并文件数目
  - 关键字
  - 合并进度
  - 输出目录

- **错误/警告提示**
  - 路径错误: 请检查提示中报错的路径是否符合上述要求
  - 关键字错误: 请检查提示中报错的文件是否含有所需的关键字 (姓名/学号/得分 etc.)
  - 未知错误: 请检查要合并的文件中是否存在多个相同关键字列, 未知错误通常由此引起, 但因为测试次数有限, 无法穷尽所有错误, 若无法自行解决, 请前往issue讨论
  - 路径警告: 若输出目录不存在或是一个文件则会在程序同一目录下输出
  - 关键字警告: 若班级最高分/平均分不存在则会警告并自动添加

---

### 4. 碎碎念

本项目经AI辅助独立完成, 未参考同类项目。如有不足请帮忙指出并提供改进方案，感激不尽。

如有任何问题欢迎在issue中提, 虽然作者大概率会直接放弃这个项目。


