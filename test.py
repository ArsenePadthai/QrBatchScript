import os
import shutil
import xlrd
import qrcode
from PIL import Image, ImageDraw, ImageFont


FONT = ImageFont.truetype('font\Roboto-Regular.ttf', size=170)
COLOR = 'rgb(0,0,0)'
OUTPUT_SIZE = (1771, 1771)


def generate_qr_label(content, dir=None):
    # dir handle
    cwd = os.getcwd()
    abs_tmp_path = os.path.join(cwd, 'tmp\%s.png' % content)
    if not os.path.exists(os.path.join(cwd, 'output\\{dir}'.format(dir=dir))):
        os.mkdir(os.path.join(cwd, 'output\\{dir}'.format(dir=dir)))
    abs_out_path = os.path.join(cwd, 'output\{0}\{1}.png'.format(dir, content))
    print(abs_tmp_path)
    print(abs_out_path)

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=20,
        border=4,
    )
    qr.add_data(content)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(abs_tmp_path)

    with Image.open(abs_tmp_path) as im:
        new_im = im.resize(OUTPUT_SIZE)
        new_im = new_im.crop((140, 190, 1631, 1681))
        new_im = new_im.resize(OUTPUT_SIZE)
        draw = ImageDraw.Draw(new_im)
        draw.text((170, 1570), content, fill=COLOR, font=FONT)
        new_im.save(abs_out_path)


if __name__ == '__main__':
    shutil.rmtree('output', ignore_errors=True)
    shutil.rmtree('tmp', ignore_errors=True)
    os.mkdir('output')
    os.mkdir('tmp')

    x1 = xlrd.open_workbook('test\\test.xlsx')
    table = x1.sheet_by_name('1000')
    for i in range(0, table.nrows):
        name = str(table.cell_value(i, 0)).split(';')[0]
        group = str(int(table.cell_value(i, 1)))
        print(name, group)
        generate_qr_label(name, group)

