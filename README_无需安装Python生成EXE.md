# 云端生成Windows EXE

1. 在GitHub新建Private repository。
2. 上传本文件夹全部内容，包括.github目录。
3. 进入Actions，运行Build Windows EXE。
4. 构建完成后，在Artifacts下载BOM_AVL采购备料一体化工具_Windows。
5. 解压得到EXE，直接分享给同事；同事无需安装Python。

本版本的总BOM usage按研发初始BOM中的Part Number顺序输出。多个BOM按文件加载顺序衔接，重复料号取首次出现位置。
