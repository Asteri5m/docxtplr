"""
Create By: Asteri5m
Create Time: 2024-01-15 11:06:12
Update Time: 2024-01-15 15:13
Version: 0.0.2
"""


import shutil
import os
from lxml import etree
import pandas as pd
from copy import copy


class myDocxTemplate:
    """针对doctpl的补充,对于特殊内容进行修改,仅支持tag模式"""

    def __init__(self, file):
        self.template_file = file

    def __del__(self):
        shutil.rmtree(self.tmp_dir)

    def renderExcel(self, file_name, chart_data):
        series = chart_data['series']
        data_frame = {'类别': chart_data['categories']}
        for data in series:
            data_frame.update(data)
        dst_df = pd.DataFrame(data_frame)
        dst_df.to_excel(file_name, index=False)
        self.render_files.append(file_name.split(self.template_file.split('.')[0])[1])

    def renderChart(self, xml_file_path, chart_data):
        tree = etree.parse(xml_file_path)
        self.namespaces = {}
        # 使用 iterwalk 遍历整个 XML 树
        for _, elem in etree.iterwalk(tree, events=("start-ns",)):
            # 提取命名空间
            name = elem[0]
            uri = elem[1]
            if name not in self.namespaces:
                self.namespaces[name] = uri

        series = chart_data['series']
        categories = chart_data['categories']

        plot_area_tag = tree.find('//c:plotArea', namespaces=self.namespaces)
        ser_tag_list = plot_area_tag.xpath('//c:ser', namespaces=self.namespaces)

        def initChartValues(tag_element, data, tag_type, index):
            if tag_type == 'cat':
                sheet_path = 'Sheet1!$A$2:$A$'
            elif tag_type == 'val':
                sheet_path = f'Sheet1!${chr(66 + index)}$2:${chr(66 + index)}$'  # ord('A') = 65
                data = list(data[index].values())[0]
            else:
                raise ValueError('unknown tag type')

            count = len(data)
            cache = None
            remove_list = []
            for element in tag_element.iter():
                if element.tag.split('}')[1] == 'f':
                    element.text = sheet_path + str(count + 1)
                if 'Cache' in element.tag.split('}')[1]:
                    cache = element
                if element.tag.split('}')[1] == 'ptCount':
                    element.set('val', str(count))
                # 删除原来的值
                if element.tag.split('}')[1] == 'pt':
                    remove_list.append(element)
            # 删除原来的值
            if cache is None:
                raise Exception('analysis xml file error: cache not find')
            for element in remove_list:
                cache.remove(element)
            # 生成新的值
            for i in range(len(data)):
                value = data[i]
                pt = etree.SubElement(cache, '{%s}pt' % self.namespaces['c'], {'idx': str(i)})
                v = etree.SubElement(pt, '{%s}v' % self.namespaces['c'])
                v.text = str(value)

        def initSer(ser_tag, index):
            f_tag          = ser_tag.find('c:f',         namespaces=self.namespaces)
            v_tag          = ser_tag.find('c:v',         namespaces=self.namespaces)
            idx_tag        = ser_tag.find('c:idx',       namespaces=self.namespaces)
            order_cat      = ser_tag.find('c:order',     namespaces=self.namespaces)
            scheme_clr_tag = ser_tag.find('a:schemeClr', namespaces=self.namespaces)

            if idx_tag is not None:
                idx_tag.set('val', str(index))
            if order_cat is not None:
                order_cat.set('val', str(index))
            if f_tag is not None:
                f_tag.text = 'Sheet1!$' + chr(ord('B') + i) + '$1'
            if v_tag is not None:
                v_tag.text = str(list(series[index].keys())[0])
            if scheme_clr_tag is not None:
                scheme_clr_tag.set('val', 'accent' + str(index + 1))

        if len(ser_tag_list) > len(series):
            parent_tag = ser_tag_list[0].getparent()
            for i in range(len(series), len(ser_tag_list)):
                parent_tag.remove(ser_tag_list[i])  # 删除多余的数据
            ser_tag_list = ser_tag_list[:len(series)]

        if len(ser_tag_list) < len(series):
            for i in range(len(ser_tag_list), len(series)):
                ser_tag = copy(ser_tag_list[0])  # 复制一个
                initSer(ser_tag, i)  # 初始化
                ser_tag_list[-1].addnext(ser_tag)  # 添加新的数据
                ser_tag_list.append(ser_tag)  # 添加到列表

        for i in range(len(ser_tag_list)):
            ser_tag = ser_tag_list[i]
            # 找到图表的列（横坐标）
            cat_tag = ser_tag.find('c:cat', namespaces=self.namespaces)
            # 找到图表的值
            val_tag = ser_tag.find('c:val', namespaces=self.namespaces)
            if cat_tag is None:
                raise Exception('analysis xml file error: cat not find')

            if val_tag is None:
                raise Exception('analysis xml file error: val not find')

            initChartValues(val_tag, series, 'val', i)
            initChartValues(cat_tag, categories, 'cat', i)

        tree.write(xml_file_path)
        self.render_files.append(xml_file_path.split(self.template_file.split('.')[0])[1])

    def renderSmartArt(self, xml_file_path, chart_data):
        """
        渲染SmartArt
        :param xml_file_path: 目标文件
        :param chart_data: 需要渲染的数据:json格式
        :return:
        """
        context = open(xml_file_path, 'r', encoding='utf-8').read()
        for key, value in chart_data.items():
            replace_str = '{{%s}}' % key
            context = context.replace(replace_str, str(value))
        open(xml_file_path, 'w', encoding='utf-8').write(context)
        self.render_files.append(xml_file_path.split(self.template_file.split('.')[0])[1])

    def initTmpDir(self):
        """
        初始化临时文件夹
        :return:
        """
        self.tmp_dir = r'tmp'

        if not os.path.exists(self.tmp_dir):
            os.mkdir(self.tmp_dir)

        self.unpack_dir = os.path.join(self.tmp_dir, self.template_file.split('.')[0])
        # 解压文件到tmp文件夹下
        shutil.unpack_archive(self.template_file, self.unpack_dir, 'zip')

    def render(self, data):
        """
        渲染模板
        :param data: 渲染数据
        :return:
        """
        self.initTmpDir()

        self.excel_dir     = r'word\embeddings'
        self.charts_dir    = r'word\charts'
        self.smart_art_dir = r'word\diagrams'
        self.render_files = []

        def analyseChartDate(chart, chart_data):
            chart_file = chart + '.xml'
            chart_path = rf'{self.tmp_dir}\{self.template_file.split(".")[0]}\{self.charts_dir}\{chart_file}'
            self.renderChart(chart_path, chart_data)

            index = int(chart[5:]) - 1
            excel_file = f'Microsoft_Excel_Worksheet{index}.xlsx'.replace('0', '')
            excel_path = rf'{self.tmp_dir}\{self.template_file.split(".")[0]}\{self.excel_dir}\{excel_file}'
            self.renderExcel(excel_path, chart_data)

        def analyseSmartArtDate(key, value):
            if len(key) > 9:
                art_files = [f'data{key[9:]}.xml']
            else:
                art_files = os.listdir(rf'{self.tmp_dir}\{self.template_file.split(".")[0]}\{self.smart_art_dir}')
                art_files = [file for file in art_files if 'data' in file]
            for art_file in art_files:
                art_path = rf'{self.tmp_dir}\{self.template_file.split(".")[0]}\{self.smart_art_dir}\{art_file}'
                self.renderSmartArt(art_path, value)

        for key, value in data.items():
            if 'chart' in key:
                analyseChartDate(key, value)
            if 'smart-art' in key:
                analyseSmartArtDate(key, value)

    def save(self, file):
        """
        保存文件
        :param file: 文件名
        :return:
        """
        file_tmp_path = os.path.join(self.tmp_dir, file)
        shutil.make_archive(file_tmp_path, 'zip', root_dir=self.unpack_dir)
        shutil.move(file_tmp_path + '.zip', file)


if __name__ == '__main__':
    template = myDocxTemplate('template.docx')

    data = {
        "smart-art": {
            "description": "替换成功",
        },
        "chart1": {
            "categories": ["类别1", "类别2", "类别3", "类别4"],
            "series": [{"系列 1": [1, 2, 3, 4]},
                       {"系列 2": [2, 3, 4, 5]},
                       {"系列 3": [3, 4, 5, 6]},
                       {"系列 4": [4, 5, 6, 7]}],
        },
    }

    template.render(data)
    template.save("result.docx")
    print(template.render_files)
