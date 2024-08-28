# pptx2pdf
批量将pptx文件转换为pdf的小工具

### 说明：
1. **`ppt_to_pdf` 函数**：此函数将单个PPT文件转换为PDF文件。使用`comtypes.client.CreateObject("PowerPoint.Application")`创建PowerPoint应用实例，并使用`presentation.SaveAs`将其保存为PDF格式。

2. **`batch_convert_ppt_to_pdf` 函数**：此函数在给定的输入文件夹中遍历所有PPT文件，并调用`ppt_to_pdf`将它们批量转换为PDF文件。

3. **输入和输出文件夹**：需要指定包含PPT文件的`input_folder`和用于保存PDF文件的`output_folder`。

### 注意：
- 该脚本只能在Windows系统上运行，因为它依赖于Windows的COM接口来调用PowerPoint应用程序。
- 需要安装Microsoft PowerPoint才能使用此脚本。
