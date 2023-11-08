# Install Rye
wget https://github.com/mitsuhiko/rye/releases/download/0.11.0/rye-x86_64-windows.exe -UseBasicParsing -OutFile rye.exe
$env:path += ";."

rye sync
rye run python -m __version__ --python --short --save

# Install WIX Toolkit
wget https://github.com/wixtoolset/wix3/releases/download/wix311rtm/wix311-binaries.zip -UseBasicParsing -OutFile c:\wix311-binaries.zip
mkdir c:\\wix311
Expand-Archive -Path c:\wix311-binaries.zip -DestinationPath c:\\wix311\\bin
$env:WIX="C:\\wix311\\"
$env:PYTHONIOENCODING="UTF-8"

# Build app
$version = @(rye run python -m __version__ --python --short)
.\pyoxidizer.exe build msi_installer --release --var version "$version"

copy-item build\\x86_64-pc-windows-msvc\\release\\msi_installer\\*.msi .\\
