# Copyright 2024 Liu Siyao
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
# Modifications Copyright 2025-2026 vanilla1108

# --coding:utf-8--
# author = ''


def convert_wmf_to_png(input_file, output_png_path):
    """
    将 WMF 数据转换为 PNG 文件。

    """
    # from PIL import ImageGrab
    # shape.Copy()
    # image = ImageGrab.grabclipboard()
    # #image.save('{}.jpg'.format(filename), 'jpeg')
    # image.save(output_png_path)

    # from PIL import Image
    # Image.open(input_file).save(output_png_path)

    from wand.image import Image

    with Image(filename=input_file) as img:
        img.format = 'png'
        img.save(filename=output_png_path)
