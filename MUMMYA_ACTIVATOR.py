import subprocess
import random
from colorama import Fore, Style

EDITIONS = {
    'home': [
        'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99',
        '3KHY7-WNT83-DGQKR-F7HPR-844BM',
        '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH',
        'PVMJN-6DFY6-9CCP6-7BKTT-D3WVR'
    ],
    'pro': [
        'W269N-WFGWX-YVC9B-4J6C9-T83GX',
        'MH37W-N47XK-V7XM9-C7227-GCQG9',
        'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
        '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
        '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
        'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'
    ],
    'education': [
        'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2',
        '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ'
    ],
    'enterprise': [
        'NPPR9-FWDCX-D2C8J-H872K-2YT43',
        'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4',
        'YYVX9-NTFWV-6MDM3-9PT4T-4M68B',
        '44RPN-FTY23-9VTTB-MP9BX-T84FV',
        'WNMTR-4C88C-JK8YV-HQ7T2-76DF9',
        '2F77B-TNFGY-69QQF-B8YKP-D69TJ',
        'DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ',
        'QFFDN-GRT3P-VKWWX-X7T3R-8B639',
        'M7XTQ-FN8P6-TTKYV-9D4CC-J462D',
        '92NFX-8DJQP-P6BBQ-THF9C-7CG2H'
    ]
}

KMS_SERVERS = [
    'kms7.MSGuides.com:1688',
    'kms8.MSGuides.com:1688',
    'kms9.MSGuides.com:1688'
]


def run_command(command):
    try:
        subprocess.check_call(command,
                              stdout=subprocess.PIPE,
                              stderr=subprocess.PIPE
                              )
    except subprocess.CalledProcessError as exc:
        raise Exception(f'Command {exc.cmd} failed: {exc.stderr.decode()}') from exc


def clear_kms_configuration():
    run_command(['cscript', '//nologo', 'slmgr.vbs', '/ckms'])
    run_command(['cscript', '//nologo', 'slmgr.vbs', '/upk'])
    run_command(['cscript', '//nologo', 'slmgr.vbs', '/cpky'])


def get_edition_keys(output):
    if 'enterprise' in output:
        return EDITIONS['enterprise']
    elif 'education' in output:
        return EDITIONS['education']
    elif 'home' in output:
        return EDITIONS['home']
    elif 'pro' in output:
        return EDITIONS['pro']
    else:
        return None


def select_kms_server():
    return random.choice(KMS_SERVERS)


def activate_windows():
    print(Fore.CYAN + 'Activating Windows...' + Style.RESET_ALL)
    clear_kms_configuration()

    output = subprocess.check_output(['wmic', 'os']).decode().lower()
    keys = get_edition_keys(output)
    if not keys:
        print(Fore.YELLOW + 'Sorry, your version is not supported.' + Style.RESET_ALL)
        return

    for key in keys:
        server = select_kms_server()
        print(Fore.GREEN + f'Trying key {key} on server {server}...' + Style.RESET_ALL)
        try:
            run_command(['cscript', '//nologo', 'slmgr.vbs', '/ipk', key])
            run_command(['cscript', '//nologo', 'slmgr.vbs', '/skms', server])
            run_command(['cscript', '//nologo', 'slmgr.vbs', '/ato'])
            print(Fore.CYAN + 'Successfully activated!' + Style.RESET_ALL)
            return
        except Exception as exc:
            print(Fore.YELLOW + f'Activation failed: {exc}' + Style.RESET_ALL)
            continue

    print(Fore.YELLOW + 'Sorry, all servers and keys have been tried, but activation failed.' + Style.RESET_ALL)


if __name__ == '__main__':
    print(Fore.CYAN + 'Activate Windows 10 (ALL versions) for FREE' + Style.RESET_ALL)
    print(Fore.GREEN + '===========================================================' + Style.RESET_ALL)
    print(
        Fore.YELLOW + '#Project: Activating Microsoft software products for FREE without additional software' + Style.RESET_ALL)
    print(Fore.GREEN + '===========================================================' + Style.RESET_ALL)
    print(Fore.YELLOW + '#Supported products:' + Style.RESET_ALL)
    print(Fore.MAGENTA + '- Windows 10 Home' + Style.RESET_ALL)
    print(Fore.MAGENTA + '- Windows 10 Professional' + Style.RESET_ALL)
    print(Fore.MAGENTA + '- Windows 10 Education' + Style.RESET_ALL)
    print(Fore.MAGENTA + '- Windows 10 Enterprise' + Style.RESET_ALL)
    print(Fore.GREEN + '===========================================================' + Style.RESET_ALL)
    print(Fore.BLUE + "Created by Black Mummya" + Style.RESET_ALL)
    activate_windows()
