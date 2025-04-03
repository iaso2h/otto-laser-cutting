import console
import config
import argparse
config.updaPath()

if not config.PARENT_DIR_PATH.exists():
    import os
    cwd = os.getcwd()
    idx = cwd.find("欧拓图纸")
    if idx > -1:
        from pathlib import Path
        config.PARENT_DIR_PATH = Path(cwd[:idx+5])
        config.updaPath()
    else:
        import sys
        print('无法找到"欧拓图纸"文件夹')
        sys.exit()



import cli

print = console.print


if __name__ == "__main__":
    argParser = argparse.ArgumentParser()
    argParser.add_argument("-L", "--legacy", action="store_true")
    argParser.add_argument("-D", "--dev",    action="store_true")
    argParser.add_argument("-R", "--rtf",    action="store_true")
    args = argParser.parse_args()
    config.DEV_MODE = args.dev
    if args.legacy:
        print(f"[bold white]版本号: {config.VERSION}[bold white]")
        print(f"[bold white]最后更新: {config.LASTUPDATED}[bold white]\n\n")
        config.GUI_MODE = False
        cli.cliStart()
        input("Press enter to proceed...")
    elif args.rtf:
        import rtfParse
        rtfParse.parseAllLog()
    else:
        config.GUI_MODE = True
        import hotkey
