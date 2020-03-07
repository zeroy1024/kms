import socket
import os
import winreg


def getInstallPath(ver):
    if ver == '1' or ver == '2' or ver == '3' or ver == '4':
        winRegKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Microsoft\\Office\\ClickToRun')
        path = winreg.QueryValueEx(winRegKey, 'InstallPath')[0] + '\\Office16'
    elif ver == '5' or ver == '6':
        winRegKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Microsoft\\Office\\15.0\\Common\\InstallRoot')
        path = winreg.QueryValueEx(winRegKey, 'Path')[0]
    elif ver == '7' or ver == '8' or ver == '9':
        winRegKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Microsoft\\Office\\14.0\\Common\\InstallRoot')
        path = winreg.QueryValueEx(winRegKey, 'Path')[0]
    return path


def serverTest():
    kmsServerList = ['kms.03k.org', 'kms.chinancce.com', 'kms.lotro.cc', 'cy2617.jios.org', 'kms.shuax.com',
                     'kms.luody.info', 'kms.cangshui.net', 'kms.library.hk', 'xykz.f3322.org', 'kms.v0v.bid',
                     'kms.moeclub.org', 'amrice.top', 'kms.lolico.moe', 'kms8.MSGuides.com', 'kms9.MSGuides.com']
    port = 1688
    for ip in kmsServerList:
        print('验证激活服务器中......')
        try:
            s = socket.socket()
            s.connect((ip, port))
            s.shutdown(2)
            print('验证成功......')
            return ip
        except socket.error as e:
            print('验证失败，更换服务器重新验证......')


def getKmsKey(ver):
    keyList = {
        '1': 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP',
        '2': '6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK',
        '3': 'XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99',
        '4': 'JNRGM-WHDWX-FJJG3-K47QV-DRTFM',
        '5': 'YC7DK-G2NP3-2QQC3-J6H88-GVGXT',
        '6': 'KBKQT-2NMXY-JJWGP-M62JB-92CD4',
        '7': 'VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB',
        '8': 'V7QKV-4XVVR-XYV4D-F7DFM-8R6BM',
        '9': 'D6QFG-VBYP2-XQHM7-J97RH-VVRCK'
    }
    return keyList[ver]


def testOspp(path):
    print(path)
    listDir = os.listdir(path)
    if 'OSPP.VBS' in listDir:
        return True
    else:
        return False


def toVol(ver, path):
    dirList = os.listdir(path + "\\..\\root\\Licenses16")
    if ver == '1' or ver == '2':
        for a in dirList:
            if "ProPlus2019VL_KMS" in a or "ProPlus2019VL_MAK" in a:
                kms_dir = path + "\\..\\root\\Licenses16\\" + a
                os.system('cscript "' + path + '\ospp.vbs" /inslic:"' + kms_dir + '"')
        input("\n\n转Vol成功，回车继续执行激活......")
    elif ver == '3' or ver == '4':
        for a in dirList:
            if "ProPlusVL_KMS" in a or "ProPlusVL_MAK" in a:
                kms_dir = path + "\\..\\root\\Licenses16\\" + a
                os.system('cscript "' + path + '\ospp.vbs" /inslic:"' + kms_dir + '"')
        input("\n\n转Vol成功，回车继续执行激活......")


if __name__ == '__main__':
    print('''
    ------------------  请选择版本  ------------------
            1. Microsoft Office 2019 专业版   
            2. Microsoft Office 2019 标准版
            3. Microsoft Office 2016 专业版
            4. Microsoft Office 2016 标准版
            5. Microsoft Office 2013 Vol 专业版
            6. Microsoft Office 2013 Vol 标准版
            7. Microsoft Office 2010 Vol 专业版
            8. Microsoft Office 2010 Vol 标准版
            9. Microsoft Office 2010 Vol 家庭版
    ------------------  请选择版本  ------------------
    ''')
    ver = input('请选择版本：')
    key = getKmsKey(ver)
    path = getInstallPath(ver)
    test = testOspp(path)
    if test:
        server = serverTest()
        if server:
            toVol(ver, path)
            print("设置密钥中")
            os.system('cscript "' + path + '\ospp.vbs" /inpkey:' + key)
            input("上述提示中是否含有 '<Product key installation successful>' 字样？ 如果有请回车继续，没有请关闭")
            os.system('cscript "' + path + '\ospp.vbs" /sethst:' + server)
            input("上述提示中是否含有 'Successfully applied setting.' 字样？ 如果有请回车继续，没有请关闭")
            os.system('cscript "' + path + '\ospp.vbs" /act')
            input("上述提示中是否含有 '<Product activation successful>' 字样？ 如果有则证明激活成功！")
    else:
        print('没有检测到OSPP.VBS，激活程序即将退出......')
        exit()
