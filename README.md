



# 常见问题解决

## 火狐浏览器中显示 折叠面板的折叠按钮 为乱码

> 这是因为火狐浏览器默认设置是禁止读取字体图标的.


1. 在火狐浏览器地址栏中输入: `about:config`
2. 搜索`security.fileuri.strict_origin_policy`, 并通过双击该项把值改为 false.

