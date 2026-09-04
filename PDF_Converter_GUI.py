"""兼容旧启动入口：不再创建 Tkinter 窗口，直接启动 Web 界面。"""

from web.app import main


if __name__ == "__main__":
    main()
