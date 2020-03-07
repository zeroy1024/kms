# kms
一款通过Python写的KMS激活工具

原理是通过注册表中的ProductName项获取当前系统版本，再根据版本从微软提供的KMS密钥中获取对应版本的密钥，然后从再验证kms激活服务器，验证成功后通过windows自带slmgr命令分别设置kms验证服务器，密钥，最后激活完成

dist中的kms.exe为Pyinstaller打包版本，可直接拿去使用，建议用于从msdn等下载的纯净镜像中，基本市面上正常的版本都可激活(Win7旗舰版除外)，Win7及以下版本(同时期Windows Server同理)需要安装微软运行库后才可使用

kms.py为源码

# Office

- [2020-03-07更新]&#160; 激活Office的工具，基本支持MSDN上的所有的Vol版本激活，2016/2019支持自动转Vol

