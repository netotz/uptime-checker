import asyncio

from colorama import init, Fore, Style
import openpyxl as opx
import aiohttp


async def check_uptime(path: str) -> None:
    print('Leyendo Excel...\n')

    red_fill = opx.styles.PatternFill(
        fill_type='solid',
        fgColor=opx.styles.colors.Color('50FF0000')
    )
    green_fill = opx.styles.PatternFill(
        fill_type='solid',
        fgColor=opx.styles.colors.Color('5000979700')
    )

    excel = opx.load_workbook(path)
    sheet = excel.active
    sheet.insert_cols(11)
    session = aiohttp.ClientSession()
    i = 0
    for row in sheet.iter_rows(min_row=8, min_col=6, max_col=11):
        if i == 10:
            break
        i += 1
        name = ' '.join(str(row[i].value) for i in range(0, 3, 1))
        if not name.isupper():
            name = 'Clasificado'

        url = row[4].value

        print(f'Contrato de {Style.BRIGHT}{name}{Style.RESET_ALL}')
        print(f'Checando enlace {Fore.CYAN}{url}{Fore.RESET}...')

        async with session.head(url) as response:
            status = response.status
            is_success = 200 <= status < 300
            color = Fore.GREEN if is_success else Fore.RED
            print(f'Código de estado: {color}{status}{Fore.RESET}\n')
            row[5].value = status

            fill = green_fill if is_success else red_fill
            row[4].fill = fill
            row[5].fill = fill

    await session.close()

    print('Guardando estados de código...')
    excel.save('estados.xlsx')


if __name__ == '__main__':
    init()
    loop = asyncio.get_event_loop()
    loop.run_until_complete(check_uptime('Z:\\honorarios.xlsx'))
