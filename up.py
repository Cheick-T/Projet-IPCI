from cx_Freeze import setup, Executable
# On appelle la fonction setup
setup(
    name = "Traitement IPCI",
    version = "1",
    description = "Un programme de traitement pour la base",
    executables = [Executable("ipci.py",icon="dcdj.ico",shortcutDir=r"C:\Users\User\Desktop")],
)
