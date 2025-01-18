import os
import random

import openpyxl
import openpyxl.cell
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment

from utils.env import env
from utils.jsondata import json_for_t2i
from utils.utils import generate_image, logger, return_x64, save_image, sleep_for_cool

README = """## 使用说明

使用前, 请打开插件目录将 "脚本文件示例.xlsx" 文件复制一份到任意位置.

然后打开复制的文件, 将要依次生成的 tag 填入 TAG 列, 保存并关闭该文件.

推文列可以不填, 为以后对接 GPT 做准备用的(实现通过推文生成 TAG 再生成图片).

回到 webui 插件页面, 设置每条 tag 生成的数量, 以及生图参数, 点击开始生成即可.

"""


if not os.path.exists("./plugins/t2i/sanp_plugin_stories_to_images/脚本文件示例.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.row_dimensions[1].height = 50
    for row in range(2, 999):
        ws.row_dimensions[row].height = 300
    for col in [f"{chr(letter)}" for letter in range(ord("A"), ord("Z") + 1)]:
        ws.column_dimensions[col].width = 40
    ws.append(
        [
            "推文",
            "TAG",
            "图片",
        ]
    )
    alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = alignment
    wb.save("./plugins/t2i/sanp_plugin_stories_to_images/脚本文件示例.xlsx")


def generate(
    excel_path,
    images_number,
    sti_negative,
    sti_width,
    sti_height,
    sti_scale,
    sti_steps,
    sti_sampler,
    sti_noise_schedule,
    sti_sm,
    sti_sm_dyn,
    sti_variety,
    sti_decrisp,
):
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook["Sheet"]
    col_num = 2
    row_num = ord("C")
    positive = sheet[f"B{col_num}"].value
    while positive is not None:
        row_num = ord("C")
        for _ in range(images_number):
            saved_path = "寄"
            while saved_path == "寄":
                json_for_t2i["input"] = positive
                json_for_t2i["parameters"]["width"] = return_x64(int(sti_width))
                json_for_t2i["parameters"]["height"] = return_x64(int(sti_height))
                json_for_t2i["parameters"]["scale"] = sti_scale
                json_for_t2i["parameters"]["sampler"] = sti_sampler
                json_for_t2i["parameters"]["steps"] = sti_steps
                if env.model != "nai-diffusion-4-curated-preview":
                    json_for_t2i["parameters"]["sm"] = sti_sm if sti_sampler != "ddim_v3" else False
                    json_for_t2i["parameters"]["sm_dyn"] = sti_sm_dyn if sti_sm and sti_sampler != "ddim_v3" else False
                    json_for_t2i["parameters"]["skip_cfg_above_sigma"] = 19 if sti_variety else None
                json_for_t2i["parameters"]["dynamic_thresholding"] = sti_decrisp
                if sti_sampler != "ddim_v3":
                    json_for_t2i["parameters"]["noise_schedule"] = sti_noise_schedule
                seed = random.randint(1000000000, 9999999999)
                json_for_t2i["parameters"]["seed"] = seed
                json_for_t2i["parameters"]["negative_prompt"] = sti_negative

                logger.debug(json_for_t2i)

                saved_path = save_image(generate_image(json_for_t2i), "t2i", seed, "None", "None")

            image = Image(saved_path)
            w = image.width
            h = image.height
            image.width, image.height = 260, int(260 / w * h)
            sheet.add_image(image, "{}{}".format(chr(row_num), col_num))

            sleep_for_cool(env.t2i_cool_time - 3, env.t2i_cool_time + 3)

            row_num += 1
        col_num += 1
        positive = sheet[f"B{col_num}"].value

    alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = alignment

    workbook.save(excel_path)
    logger.success("处理完成!")
    return "处理完成!"
