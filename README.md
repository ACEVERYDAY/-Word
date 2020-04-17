## 前言
亲人工作考试，公司给的题库好像是直接从数据库导出的表格Excel形式，在移动端上非常难看，需要不断左右上下滑动，看不了多少题眼就瞎了，遂主动请缨编写python脚本解决之。
原本给的题库在手机上横屏显示是这样的↓↓↓（想象一下是在手机上）无比恶心
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200412192742698.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2NwcmltZXNwbHVz,size_16,color_FFFFFF,t_70)

## 我的工作
公司给出的格式是.xlsx的（Excel表格的默认格式），盲猜是直接从答题数据库导出的，表名和属性名应该是稍微做了从英文到中文的改变，然后，就直接这样发给员工了....
表格有八个，放在一个文件夹下，由于不同工种的题表头是相同的，因此可以编写代码统一处理。
首先是获取题库存放路径，便于对指定路径文件处理：

```python
to_path = r'D:\python_project\TableAfterProcessing'
dir_path = r'D:\python_project\题库名\backup'
name_list = []
for i in os.listdir(dir_path):
    name_list.append(i)
```

之前学过python库pandas的基本操作，由于一个月前数模美赛的时候使用过并使用博客记录，因此总体来说还不算生疏。
关于Excel表格的读取，作者首先手动将表格转换成了.csv格式（表格不多，因此没必要编写代码了，当然，如果愿意还是可以的）。
观察到表格中知识点一栏数据完全相同，选项个数一栏并没有什么参考价值，因此去掉这两行，只保留题型，题干，选项，答案。

然后就是采用pandas将缺失值null变为空字符串' '，这样的目的是避免将null这个字符写入到word。

```python
for file_name in name_list:
    # 文件读取路径
    from_path = os.path.join(dir_path, file_name)
    p_data = pd.read_csv(from_path, engine='python', usecols=['题型', '题干', '选项', '答案'])
    p_data = p_data.where(p_data.notnull(), '')
```
经过对数据的处理后预处理后，表格便只剩下了四列数据，清爽了很多。
然而光是这样还是不够的，毕竟涉及到手机端浏览表格就得放大，滑动，一不小心点到格子里去还要点出来，对用户很不友好。
因此，我决定将表格数据导入到word，变成常见的题型格式。
这就需要用到python的docx库，关于这个库的讲解就不在这里赘述了，笔者也是通过百度新学习的，这里主要说一下设计和逻辑。

**1.题型归类**
题型分为单选题，多选题，判断题。表格中对于每一个题都有其对应的类型描述，无外这三种。同时，同一类的数据是聚集在一起的，因此，可以设置标志位记录前一个题目所属的题型，如果当前类别和上一个相同，则只需要写入题号题干等；如果不同，就使用docx中的`Document.add_heading()`方法新建立一个标题。

**2.正确答案标红**
如果单纯的将答案写在每一个题的后面或者开头，这样固然可以，但显然不够直观。一种友好的方式是将正确答案标为红色，这样便能直观的看出。
如何实现呢？
原本表格中的答案是以'ABC'这样的方式给出的，python中自带关键字`in `可以用来判断A串是否连续存在于B中，例如`'as' in 'asda'`，返回值是`True`，而`'sa' in 'asda'`返回值则是`False`。
故而拿到了选项后，只需要使用`str.split()`方法切分字符串，再依次判断每个字符串的首个字符是否存在于正确答案字符串中就可以了。
拿这组数据举例：
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200412201556172.png)
给定选项有：A.劳动生产率　　B.产品质量　　   C.产量　　 D.工作质量
因此切分后的字符串列表是这样的`['A.劳动生产率', 'B.产品质量', 'C.产量',' D.工作质量']`
正确答案字符串为`'A,B,D'`
取其中第一个字符串`'A.劳动生产率'`，首个字符为`'A'`，A存在于'A,B,D'中，证明这条答案是正确的，因此调用docx库自带的方法将字符串写入到word并标记为红色。

```python
# 若为判断题则将答案写入
if q_type == '判断题':
    document.add_paragraph(u'答案：' + str(r_ans) + '\n')
# 否则只标红正确选项
else:
    res_list = (str(r_choose)).split()
    # print(res_list)
    p.add_run('\n')
    for res in res_list:
        run = p.add_run(str(res) + ' ')
        # print(res[0])
        if res[0] in r_ans:
            run.font.color.rgb = RGBColor(255,0,0)
    p.add_run('\n')
```

经过我一通操作后变成了这样↓↓↓
单选题
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200412202946681.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2NwcmltZXNwbHVz,size_16,color_FFFFFF,t_70)
多选题
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200412194437253.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2NwcmltZXNwbHVz,size_16,color_FFFFFF,t_70)
判断题
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200412202838825.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2NwcmltZXNwbHVz,size_16,color_FFFFFF,t_70)

## 代码
这里放上整个代码，若有需要的同学可以作为参考。

```python
# *-* encoding:utf-8 *-*
import os
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor


to_path = r'D:\python_project\TableAfterProcessing'
dir_path = r'D:\python_project\XXXX\backup'
name_list = []
for i in os.listdir(dir_path):
    name_list.append(i)

print(name_list)
# ['D:\\python_project\\XXX考试题库\\backup\\ssss.csv']

for file_name in name_list:
    # 文件读取路径
    from_path = os.path.join(dir_path, file_name)
    # 创建文档对象，设置字体
    document = Document()
    document.styles['Normal'].font.name = u'宋体'
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    print((file_name))
    p_data = pd.read_csv(from_path, engine='python', usecols=['题型', '题干', '选项', '答案'])
    p_data = p_data.where(p_data.notnull(), '')
    # 读取指定几列
    q_type = ''
    for index in range(len(p_data)):
        r_type = p_data['题型'][index]
        r_cont = p_data['题干'][index]
        r_choose = p_data['选项'][index]
        r_ans = p_data['答案'][index]
        # print(str(r_choose))
        # 判断当前题型，确定是否创建对应类别标题
        if r_type is not q_type:
            q_type = r_type
            document.add_heading(q_type)
        # 将题号以及题干写入文档
        p = document.add_paragraph(str(index+1) + r'.' + str(r_cont))
        # 写入选项
        # if str(r_choose).strip() is not '':
        #     document.add_paragraph(str(r_choose))
        # 若为判断题则将答案写入
        if q_type == '判断题':
            document.add_paragraph(u'答案：' + str(r_ans) + '\n')
        # 否则只标红正确选项
        else:
            res_list = (str(r_choose)).split()
            # print(res_list)
            p.add_run('\n')
            for res in res_list:
                run = p.add_run(str(res) + ' ')
                # print(res[0])
                if res[0] in r_ans:
                    run.font.color.rgb = RGBColor(255,0,0)
            p.add_run('\n')
            # p.add_run('')
            # 切分答案字符串
        # pass
    # 写入对应路径
    document.save(os.path.join(to_path, file_name[0:-4]+'.docx'))


```

