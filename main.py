from CrystalDex import CrystalDex_main

if __name__ == '__main__':
    app = CrystalDex_main()
    app.startup()
    app.root.after_idle(app.refocus)
    app.root.mainloop()
    app.Server_Save()