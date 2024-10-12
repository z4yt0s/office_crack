def main() -> None:
    banner = '''
________   _____  _____.__             _________                       __
\\_____  \\_/ ____\\/ ____\\__| ____  ____ \\_   ___ \\____________    ____ |  | __
 /   |   \\   __\\\\   __\\|  |/ ___\\/ __ \\/    \\  \\/\\_  __ \\__  \\ _/ ___\\|  |/ /
/    |    \\  |   |  |  |  \\  \\__\\  ___/\\     \\____|  | \\// __ \\\\  \\___|    <
\\_______  /__|   |__|  |__|\\___  >___  >\\______  /|__|  (____  /\\___  >__|_ \\
        \\/                     \\/    \\/        \\/            \\/     \\/     \\/ '''
    print(banner)

    URL: str = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe'
    current_path: Path = Path(__file__).parent

    filename: str = URL.split("/")[-1]
    tmp_dir: Path = current_path / 'tmp'
    office_dir : Path = current_path / 'office'
    setup_file = tmp_dir / filename

    print(f'[*] Creando carpetas tmp/ & office/')
    tmp_dir.mkdir(exist_ok=True)
    office_dir.mkdir(exist_ok=True)

    response = get(URL, stream=True)
    if not response.status_code == 200:
        print('[!] Existe un problema al descargar el archivo')
        exit(0)

    print('[*] Descargando archivos requeridos')
    with open(setup_file, 'wb') as file:
        for chunk in response.iter_content(chunk_size=8192):
            if chunk: file.write(chunk)
    print(f'[*] Archivo descargado y guardado en: {setup_file}')

    try:
        print(f'[*] Ejecutando el archivo {setup_file}')
        run([str(setup_file), f'/extract:{str(office_dir)}'], check=True)
    except CalledProcessError as e:
        print(f'[!] Error al ejecutar el archivo: {e}')

    file_del = office_dir / 'configuration-Office365-x64.xml'
    file_del.unlink()

    setup_exe: Path = office_dir / 'setup.exe'
    config_xml: Path = office_dir / 'configuration.xml'
    try:
        print(f'[*] Ejecutando el archivo {setup_exe}')
        run([str(setup_exe), '/configure', str(config_xml)], check=True)
    except CalledProcessError as e:
        print(f'[!] Error al ejecutar el archivo: {e}')

    print('[*] Paquete Office pirateado correctamente')

if __name__ == '__main__':
    main()
