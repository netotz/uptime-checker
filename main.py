import asyncio

from colorama import init, Fore, Style
import pandas as pd
import aiohttp


async def check_uptime(path: str) -> None:
    print('Leyendo Excel...\n')
    excel = pd.read_excel(
        path,
        engine='openpyxl',
        header=6,
        usecols=[5, 6, 7, 9]
    )

    statuses = list()
    session = aiohttp.ClientSession()
    for index, row in excel.iterrows():
        name = ' '.join(str(row[i]) for i in range(5, 8, 1))
        if not name.isupper():
            name = 'Clasificado'
        url = row[9]

        print(f'Contrato de {Style.BRIGHT}{name}{Style.RESET_ALL}')
        print(f'Checando enlace {Fore.CYAN}{url}{Fore.RESET}...')

        async with session.head(url) as response:
            status = response.status
            color = Fore.GREEN if 200 <= status < 300 else Fore.RED
            print(f'Código de estado: {color}{status}{Fore.RESET}\n')
            statuses.append(status)
    await session.close()

    excel.insert(10, 'Estado del hipervínculo al contrato', statuses)

    print('Guardando Excel...')
    excel.to_excel('estados.xlsx', engine='openpyxl')


if __name__ == '__main__':
    init()
    loop = asyncio.get_event_loop()
    loop.run_until_complete(check_uptime('Z:\\honorarios.xlsx'))
