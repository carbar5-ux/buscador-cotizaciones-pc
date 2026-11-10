; Instalador de Windows del Buscador de cotizaciones (Inno Setup 6).
; Lo arma GitHub Actions (.github/workflows/instalador.yml) después de instalador/armar.ps1.
; Se instala en la carpeta de programas del usuario (sin permisos de administrador), para que el programa
; pueda seguir actualizándose solo desde GitHub.

#ifndef Version
  #define Version "0.0.0"
#endif

[Setup]
AppId={{8C2F6E1A-4B7D-4E3A-9A51-6B0C2D7E9F13}
AppName=Buscador de cotizaciones
AppVersion={#Version}
AppVerName=Buscador de cotizaciones {#Version}
AppPublisher=Carlos
AppPublisherURL=https://github.com/carbar5-ux/buscador-cotizaciones-pc
DefaultDirName={localappdata}\Programs\Buscador de cotizaciones
DefaultGroupName=Buscador de cotizaciones
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=salida
OutputBaseFilename=Instalar-Buscador-Cotizaciones-{#Version}
SetupIconFile=..\icono_cotizaciones.ico
UninstallDisplayIcon={app}\icono_cotizaciones.ico
UninstallDisplayName=Buscador de cotizaciones
LicenseFile=..\LICENSE
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "escritorio"; Description: "Crear accesos directos en el Escritorio"
Name: "alencender"; Description: "Abrir el buscador al encender el PC (minimizado)"; Flags: unchecked

[Files]
Source: "app\*"; DestDir: "{app}"; Excludes: "config.toml"; Flags: ignoreversion recursesubdirs createallsubdirs
; La configuración solo se instala la primera vez: reinstalar o actualizar nunca la cambia ni la borra
Source: "app\config.toml"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall

[Icons]
Name: "{autoprograms}\Buscador de cotizaciones\Buscador Cotizaciones"; Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\ventana.py"""; WorkingDir: "{app}"; IconFilename: "{app}\icono_cotizaciones.ico"
Name: "{autoprograms}\Buscador de cotizaciones\Excel Cotizaciones"; Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\buscador_cotizaciones.py"" --abrir"; WorkingDir: "{app}"; IconFilename: "{app}\icono_excel.ico"
Name: "{autoprograms}\Buscador de cotizaciones\Manual de uso"; Filename: "{app}\Manual_de_uso.pdf"
Name: "{autodesktop}\Buscador Cotizaciones"; Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\ventana.py"""; WorkingDir: "{app}"; IconFilename: "{app}\icono_cotizaciones.ico"; Tasks: escritorio
Name: "{autodesktop}\Excel Cotizaciones"; Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\buscador_cotizaciones.py"" --abrir"; WorkingDir: "{app}"; IconFilename: "{app}\icono_excel.ico"; Tasks: escritorio
Name: "{autodesktop}\Manual Cotizaciones"; Filename: "{app}\Manual_de_uso.pdf"; Tasks: escritorio
Name: "{userstartup}\Buscador de cotizaciones"; Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\ventana.py"" --minimizado"; WorkingDir: "{app}"; IconFilename: "{app}\icono_cotizaciones.ico"; Tasks: alencender

[Run]
Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\ventana.py"""; WorkingDir: "{app}"; Description: "Abrir el Buscador de cotizaciones ahora"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Lo que crea el programa mismo (no se tocan el Excel ni las cotizaciones, que están en otra carpeta)
Type: filesandordirs; Name: "{app}\__pycache__"
Type: filesandordirs; Name: "{app}\python"
Type: files; Name: "{app}\registro.log"
