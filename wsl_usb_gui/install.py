import appdirs
import logging
import subprocess
import sys
import time
import wx
import re
from pathlib import Path

user_data_dir = Path(appdirs.user_data_dir("wsl-usb-gui", ""))
user_data_dir.mkdir(parents=True, exist_ok=True)

install_log = user_data_dir / "install.log"
print("日志记录到", install_log)
logging.basicConfig(format="%(asctime)s | %(levelname)-8s | %(message)s", filename=install_log, encoding='utf-8', level=logging.DEBUG)

logging.info("正在运行安装后脚本")


def run(args, show=False):
    CREATE_NO_WINDOW = 0x08000000
    return subprocess.run(
        args,
        capture_output=True,
        creationflags=0 if show else CREATE_NO_WINDOW,
        shell=(isinstance(args, str))
    )

def msgbox_ok_cancel(title, message):
    ret = wx.MessageDialog(
        parent=None,
        caption=title,
        message=message,
        style=wx.OK | wx.CANCEL | wx.ICON_INFORMATION,
    )
    return ret == wx.ID_OK


# Check WSL Version
def check_wsl_version():
    try:
        wsl_installs_ret = run("wsl --list -v").stdout
        wsl_installs = wsl_installs_ret.decode().replace("\x00", "").split("\n")
        iname = wsl_installs[0].find("NAME")
        istate = wsl_installs[0].find("STATE")
        iversion = wsl_installs[0].find("VERSION")
        if -1 in (iname, istate, iversion):
            logging.error("警告：无法检查WSL2版本")
        else:
            for row in wsl_installs[1:]:
                if row.strip().startswith("*"):
                    name = row[iname:istate].strip()
                    version = row[iversion:].strip()
                    if version != "2":
                        logging.warning(f"默认WSL ({name}) 需要更新到版本2")
                        update_wsl_version(name)
                    else:
                        logging.info(f"默认WSL ({name}) 已是版本2")
                    break
        return True
    except Exception as ex:
        logging.exception(ex)
        rsp = msgbox_ok_cancel(
            title="检查WSL版本时出错",
            message=f"检查WSL版本时发生意外错误，是否继续？",
        )
        if not rsp:
            raise SystemExit(ex)
    return False


def update_wsl_version(name):
    try:
        rsp = msgbox_ok_cancel(
                title="WSL转换为版本2？",
                message=f"默认WSL ({name}) 需要更新到版本2，现在进行吗？",
            )
        if rsp:
            run(f'wsl --set-version "{name}" 2', show=True)
        return True
    except Exception as ex:
        logging.exception(ex)
        rsp = msgbox_ok_cancel(
            title="转换WSL版本时出错",
            message=f"转换WSL到版本2时发生意外错误，是否继续？",
        )
        if not rsp:
            raise SystemExit(ex)
    return False


def check_kernel_version():
    try:
        version = run(["C:\Windows\System32\wsl.exe", "--", "/bin/uname", "-r"]).stdout.decode().strip()
        logging.info(f"WSL2 内核: {version}")
        number = version.split("-")[0]
        number_tuple = tuple((int(n) for n in number.split(".")))
        if number_tuple < (5,10,60,1):
            logging.warning("内核需要更新")
            run(r'''Powershell -Command "& { Start-Process \"wsl\" -ArgumentList @(\"--update\") -Verb RunAs } "''')
            run("wsl --shutdown")
        return True

    except Exception as ex:
        logging.exception(ex)
        rsp = msgbox_ok_cancel(
            title="检查或升级WSL内核版本时出错",
            message=f"检查或升级WSL内核版本时发生意外错误，是否继续？",
        )
        if not rsp:
            raise SystemExit(ex)
    return False

def wsl_sudo_run(command):
    return run(["wsl", "--user", "root", "bash", "-c", command]).stdout.strip().decode()


def install_client():
    try:
        logging.info("正在安装WSL客户端工具：")
        logging.info(wsl_sudo_run("apt install -y linux-tools-generic hwdata"))
        latest = wsl_sudo_run("ls -vr1 /usr/lib/linux-tools/*/usbip | head -1")
        logging.info(f"Latest version installed: {latest}")
        logging.info(wsl_sudo_run(f"update-alternatives --verbose --install /usr/local/bin/usbip usbip \"{latest}\" 20"))
        return True
    except Exception as ex:
        logging.exception(ex)
        rsp = msgbox_ok_cancel(
            title="安装Linux工具时出错",
            message=f"安装Linux工具时发生意外错误，是否继续？",
        )
        if not rsp:
            raise SystemExit(ex)
    return False




def find_installers():
    msi = None
    msi_vers = (0, 0, 0)
    try:
        src_dir = Path(__file__).parent.parent.resolve()
    except:
        src_dir = Path(".")
    installers = list(app_dir.glob("usbipd-win*.msi")) + list(src_dir.glob("usbipd-win*.msi"))
    installers.sort()
    if installers:
        msi = installers[0]
        try:
            msi_vers = tuple((int(v) for v in (re.findall(r'_(\d+\.\d+\.\d+)_', msi.name)[0].split("."))))
        except:
            pass
    return msi, msi_vers


app_dir = Path(sys.executable).parent.resolve()
MSI, MSI_VERS = find_installers()


def install_server():

    try:
        if not MSI:
            msg = f"无法在以下位置找到usbipd-win安装程序： {app_dir}"
            raise OSError(msg)

        usbipd_install_log = user_data_dir / "usbipd_install.log"
        cmd = f'msiexec /i "{MSI}" /passive /norestart /log "{usbipd_install_log}"'
        logging.info(cmd.encode())
        rsp = run(cmd)
        return True
    except Exception as ex:
        logging.exception(ex)
        rsp = msgbox_ok_cancel(
            title="安装usbipd-win时出错",
            message=f"安装Windows服务器时发生意外错误，是否继续？",
        )
        if not rsp:
            raise SystemExit(ex)
    return False


root_win = None


def install_task(parent=None):
    if parent is None:
        __app = wx.App()

    progress_bar = wx.ProgressDialog(parent=parent, message='正在安装...',
                       title='正在安装', maximum=4)

    # Install tasks
    progress_bar.Update(0, "检查/更新WSL到WSL2...")
    rsp = check_wsl_version()

    progress_bar.Update(1, "检查/更新WSL2内核版本...")
    rsp &= check_kernel_version()

    progress_bar.Update(3, "安装usbipd-win服务...")
    rsp &= install_server()

    progress_bar.Update(4, "已完成。")
    time.sleep(1)

    logging.info("已完成")
    progress_bar.Destroy()
    return rsp


if __name__ == "__main__":
    install_task()
