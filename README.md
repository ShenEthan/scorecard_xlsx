# scorecard_xmlx文档说明

基于***scorecardpy***，将中间结果整理成文档写入Excel中。


## p01_data_prepare
>文件p01_data_prepare.py，基于scorecardpy产生包括数据的分箱，iv计算，psi计算等；

## p03_result_xlsx_create
>文件p03_result_xlsx_create.py，基于产生的结果数据，写入excel结果文档。

`import pandas as pd`


```
import pandas as pd
from pandas.api.types import is_string_dtype
from pandas.api.types import is_numeric_dtype
import numpy as np
import scorecardpy as sc
import xlsxwriter as xl
import sys
```

![WechatIMG184.jpeg](https://i.loli.net/2020/10/13/s89J7IrLZ1jdDpt.jpg)

|  表头   | 表头  |
|  ----  | ----  |
| 单元格  | 单元格 |
| 单元格  | 单元格 |