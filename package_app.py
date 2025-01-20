# package_app.py
import os
import sys
import shutil
import subprocess
from pathlib import Path
import PyInstaller.__main__
import logging

class MDMProcessorPackager:
    def __init__(self):
        self.root_dir = Path(__file__).parent
        self.build_dir = self.root_dir / 'build'
        self.dist_dir = self.root_dir / 'dist'
        self.config_dir = self.root_dir / 'config'
        self.src_dir = self.root_dir / 'src'
        self.install_dir = self.root_dir / 'install'
        self.logs_dir = self.root_dir / 'logs'

        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler('packaging.log')
            ]
        )
        self.logger = logging.getLogger(__name__)

    def create_packaging_environment(self):
        """Install required packaging tools"""
        try:
            requirements = [
                'pyinstaller',
                'requests',
                'office365',
                'pandas',
                'quickbase_client'
            ]
            for req in requirements:
                self.logger.info(f"Installing {req}...")
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', req])
        except Exception as e:
            self.logger.error(f"Error installing requirements: {str(e)}")
            raise

    def create_version_file(self):
        """Create version info file"""
        try:
            version_file = self.build_dir / 'version_info.txt'
            version_info = '''
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'Wesco'),
         StringStruct(u'FileDescription', u'MDM Processor'),
         StringStruct(u'FileVersion', u'1.0.0'),
         StringStruct(u'InternalName', u'mdm_processor'),
         StringStruct(u'LegalCopyright', u'Copyright (c) 2024 Wesco'),
         StringStruct(u'OriginalFilename', u'MDM_Processor.exe'),
         StringStruct(u'ProductName', u'MDM Processor'),
         StringStruct(u'ProductVersion', u'1.0.0')])
    ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
'''
            version_file.write_text(version_info)
            self.logger.info(f"Created version file at {version_file}")
            return version_file
        except Exception as e:
            self.logger.error(f"Error creating version file: {str(e)}")
            raise

    def create_spec_file(self):
        """Create PyInstaller spec file"""
        try:
            spec_file = self.build_dir / 'mdm_processor.spec'
            
            # Find the main UI script
            ui_script = None
            for file in self.src_dir.glob('*.py'):
                if 'mdm_processor_ui' in file.name.lower():
                    ui_script = file
                    break
            
            if not ui_script:
                raise FileNotFoundError("Could not find main UI script")

            # Get all MDM processor scripts
            mdm_scripts = []
            for file in self.src_dir.glob('*_mdm_processor.py'):
                if file != ui_script:
                    mdm_scripts.append(file)

            # Create datas list for additional files
            datas = []
            if self.config_dir.exists():
                datas.extend([
                    (str(f), str(f.relative_to(self.root_dir).parent))
                    for f in self.config_dir.glob('*')
                ])
            
            for script in mdm_scripts:
                datas.append((str(script), str(script.relative_to(self.root_dir).parent)))

            datas_str = ',\n                '.join(
                f"(r'{src}', '{dst}')" for src, dst in datas
            )

            spec_content = f'''
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    [r'{str(ui_script)}'],
    pathex=[r'{str(self.root_dir)}'],
    binaries=[],
    datas=[
        {datas_str}
    ],
    hiddenimports=['pandas', 'office365', 'quickbase_client'],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MDM_Processor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
    icon=None,
)
'''
            spec_file.write_text(spec_content)
            self.logger.info(f"Created spec file at {spec_file}")
            return spec_file
        except Exception as e:
            self.logger.error(f"Error creating spec file: {str(e)}")
            raise

    def create_installer_batch(self):
        """Create installer batch file"""
        try:
            install_file = self.install_dir / 'install.bat'
            batch_content = '''@echo off
echo Installing MDM Processor...

echo Creating directories...
mkdir "%PROGRAMFILES%\\Wesco\\MDM Processor" 2>nul
mkdir "%PROGRAMFILES%\\Wesco\\MDM Processor\\config" 2>nul
mkdir "%PROGRAMFILES%\\Wesco\\MDM Processor\\logs" 2>nul

echo Copying files...
xcopy /Y /E /I "*.exe" "%PROGRAMFILES%\\Wesco\\MDM Processor\\"
xcopy /Y /E /I "config\\*" "%PROGRAMFILES%\\Wesco\\MDM Processor\\config\\"

echo Creating shortcuts...
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\\Desktop\\MDM Processor.lnk');$s.TargetPath='%PROGRAMFILES%\\Wesco\\MDM Processor\\MDM_Processor.exe';$s.Save()"
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%PROGRAMDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\MDM Processor.lnk');$s.TargetPath='%PROGRAMFILES%\\Wesco\\MDM Processor\\MDM_Processor.exe';$s.Save()"

echo Installation complete!
pause
'''
            install_file.write_text(batch_content)
            self.logger.info(f"Created installer batch file at {install_file}")
            return install_file
        except Exception as e:
            self.logger.error(f"Error creating installer batch file: {str(e)}")
            raise

    def run_pyinstaller(self, spec_file):
        """Run PyInstaller using Python module"""
        try:
            self.logger.info("Running PyInstaller...")
            PyInstaller.__main__.run([
                str(spec_file)
            ])
            self.logger.info("PyInstaller completed successfully")
        except Exception as e:
            self.logger.error(f"PyInstaller error: {str(e)}")
            raise

    def package_application(self):
        """Main packaging function"""
        try:
            self.logger.info("Starting packaging process...")
            
            # Create necessary directories
            for directory in [self.build_dir, self.dist_dir, self.install_dir]:
                directory.mkdir(exist_ok=True)
                self.logger.info(f"Created directory: {directory}")
            
            # Create required files
            version_file = self.create_version_file()
            spec_file = self.create_spec_file()
            install_file = self.create_installer_batch()
            
            # Run PyInstaller
            self.run_pyinstaller(spec_file)
            
            # Create installation package
            self.logger.info("Creating installation package...")
            install_pkg = self.dist_dir / 'MDM_Processor_Install'
            install_pkg.mkdir(exist_ok=True)
            
            # Copy files
            if (self.dist_dir / 'MDM_Processor.exe').exists():
                shutil.copy(self.dist_dir / 'MDM_Processor.exe', install_pkg)
                self.logger.info("Copied executable")
            
            if self.config_dir.exists():
                shutil.copytree(self.config_dir, install_pkg / 'config', dirs_exist_ok=True)
                self.logger.info("Copied configuration files")
            
            shutil.copy(install_file, install_pkg)
            self.logger.info("Copied installer batch file")
            
            # Create README
            readme_content = '''MDM Processor Application
Version 1.0.0

Installation Instructions:
1. Run install.bat as administrator
2. The application will be installed to Program Files
3. Shortcuts will be created on Desktop and Start Menu

For support, contact IT Support

Copyright © 2024 Wesco'''
            
            (install_pkg / 'README.txt').write_text(readme_content)
            self.logger.info("Created README file")
            
            # Create ZIP archive
            zip_path = str(self.dist_dir / 'MDM_Processor_Install')
            shutil.make_archive(zip_path, 'zip', install_pkg)
            self.logger.info(f"Created installation package: {zip_path}.zip")
            
            self.logger.info("Packaging complete!")
            
        except Exception as e:
            self.logger.error(f"Packaging failed: {str(e)}")
            raise

def main():
    try:
        packager = MDMProcessorPackager()
        packager.create_packaging_environment()
        packager.package_application()
    except Exception as e:
        print(f"Error during packaging: {str(e)}")
        print("Check packaging.log for details")
        sys.exit(1)

if __name__ == '__main__':
    main()