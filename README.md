# Excel拼音

用于在Excel中将汉字换为拼音的宏

## 使用

将这个宏加入Excel

使用函数getpy将字符串转化为拼音

使用函数getszm将字符串转化为拼音首字母

replaceWord参数控制未知字符的替换值

emptyWord参数控制空字符的返回值

拼音取常用拼音，多音字可能会出现错误

拼音之间使用空格分隔

### e.g.

=getpy("这里是宽宽的作品,我是宽宽","^")

->kuan kuan de zuo pin ^ wo shi kuan kuan

=getpy("这里是宽宽的作品,我是宽宽")

->kuan kuan de zuo pin ? wo shi kuan kuan

=getszm("宽宽的作品")

->kkdzp

=getszm("","?","空")

->空

## 关于作者

作者主页[宽宽2007](https://kuankuan2007.gitee.io "作者主页")

本项目在[苟浩铭/Excel拼音 (gitee.com)](https://gitee.com/kuankuan2007/excel-pinyin)上开源
