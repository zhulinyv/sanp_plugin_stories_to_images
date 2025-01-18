from pathlib import Path

import gradio as gr

from plugins.t2i.sanp_plugin_stories_to_images.utils import README, generate
from src.text2image_nsfw import return_resolution
from utils.env import env
from utils.utils import NOISE_SCHEDULE, RESOLUTION, SAMPLER, open_folder, read_json

webui_language = read_json(f"./files/languages/{env.webui_lang}/webui.json")


def plugin():
    with gr.Tab("推文生图工具"):
        open_plugin_directory = gr.Button("打开插件目录")
        generate_button = gr.Button("开始生成")
        excel_path = gr.Textbox(label="Excel 文件路径")
        images_number = gr.Slider(1, 20, 5, step=1, label="每段 tag 生成图片的数量")
        gr.Markdown("<hr>")
        sti_negative = gr.Textbox("<negative:默认负面提示词>,", label="负面提示词", lines=3)
        with gr.Row():
            sti_resolution = gr.Dropdown(
                RESOLUTION,
                value=("832x1216" if env.img_size == -1 else "{}x{}".format((env.img_size)[0], (env.img_size)[1])),
                label=webui_language["t2i"]["resolution"],
            )
            sti_width = gr.Textbox(
                value=(env.img_size)[0] if env.img_size != -1 else "832",
                label=webui_language["t2i"]["width"],
            )
            sti_height = gr.Textbox(
                value=(env.img_size)[1] if env.img_size != -1 else "1216",
                label=webui_language["t2i"]["height"],
            )
            sti_resolution.change(
                return_resolution,
                sti_resolution,
                outputs=[sti_width, sti_height],
                show_progress="hidden",
            )
        with gr.Row():
            sti_scale = gr.Slider(
                minimum=0,
                maximum=10,
                value=env.scale,
                step=0.1,
                label=webui_language["t2i"]["scale"],
            )
            sti_steps = gr.Slider(minimum=0, maximum=50, value=env.steps, step=1, label=webui_language["t2i"]["steps"])
        with gr.Row():
            sti_sampler = gr.Dropdown(
                SAMPLER,
                value=env.sampler,
                label=webui_language["t2i"]["sampler"],
            )
            sti_noise_schedule = gr.Dropdown(
                NOISE_SCHEDULE,
                value=env.noise_schedule,
                label=webui_language["t2i"]["noise_schedule"],
            )
        with gr.Row():
            sti_sm = gr.Checkbox(
                value=env.sm if env.model != "nai-diffusion-4-curated-preview" else False,
                label="sm",
                visible=True if env.model != "nai-diffusion-4-curated-preview" else False,
            )
            sti_sm_dyn = gr.Checkbox(
                value=env.sm_dyn if env.model != "nai-diffusion-4-curated-preview" else False,
                label=webui_language["t2i"]["smdyn"],
                visible=True if env.model != "nai-diffusion-4-curated-preview" else False,
            )
            sti_variety = gr.Checkbox(
                value=env.variety if env.model != "nai-diffusion-4-curated-preview" else False,
                label="variety",
                visible=True if env.model != "nai-diffusion-4-curated-preview" else False,
            )
            sti_decrisp = gr.Checkbox(value=env.decrisp, label="decrisp")
        opt_info = gr.Textbox(label="输出信息")
        gr.Markdown(README)
        open_plugin_directory.click(
            open_folder, inputs=gr.Textbox(Path("./plugins/t2i/sanp_plugin_stories_to_images"), visible=False)
        )
        generate_button.click(
            generate,
            inputs=[
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
            ],
            outputs=opt_info,
        )
