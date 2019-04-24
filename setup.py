import os, sys
import imp
import winreg
from win32com.client import Dispatch
from setuptools import setup
from setuptools.command.install import install as _install


__package__ = 'pyxcel'
__version__ = '1.0.2'


here = os.path.dirname(os.path.abspath(__file__))


with open(os.path.join(here, "requirements.txt")) as f:
    required = f.read().splitlines()


with open(os.path.join(here, "README.md")) as f:
    __doc__ = f.read()


def createShortcut(path, target='', wDir='', icon=''):
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wDir
    shortcut.IconLocation = icon
    shortcut.save()


def get_reg(name, path):
    # Read variable from Windows Registry
    # From https://stackoverflow.com/a/35286642
    try:
        registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0,
                                       winreg.KEY_READ)
        value, regtype = winreg.QueryValueEx(registry_key, name)
        winreg.CloseKey(registry_key)
        return value
    except WindowsError:
        return None


def post_install():
    _, pkg_folder, _ = imp.find_module(__package__, sys.path[1:])
    desktop_folder = os.path.join(os.environ["userprofile"], "Desktop")
    wDir = os.path.join(pkg_folder, "launcher")
    target = os.path.join(wDir, __package__ + ".bat")
    icon = os.path.join(wDir, __package__ + ".ico")
    lnkPath = os.path.join(desktop_folder, __package__ + ".lnk")
    print(lnkPath, target, wDir, icon)
    createShortcut(lnkPath, target, wDir, icon)



class install(_install):
    def run(self):
        _install.run(self)
        self.execute(post_install, (),
                     msg="Running post install task")


setup(
    name=__package__,
    version=__version__,
    author='joao ceia',
    author_email='joao.p.ceia@gmail.com',
    packages=['pyxcel', 'pyxcel.launcher', 'pyxcel.vba', 'pyxcel.examples'],
    cmdclass={'install': install},
    url='https://github.com/jpceia/pyxcel',
    description=__doc__,
    install_requires=required,
    python_requires='>=3.6',
    platforms=['Windows'],
    package_data={
        'pyxcel.launcher': ['pyxcel.bat', 'pyxcel.ico'],
        'pyxcel.vba': ['pyxcel.xla']},
    include_package_data=True,
    zip_safe=False,
    classifiers=[
        'Programming Language :: Python :: 3',
        'Operating System :: Microsoft :: Windows',
    ],
)
