import pandas as pd
import yaml
from yaml import Dumper

class OrderedDumper(Dumper):
        pass

def represent_dict_order(self, data):
    return self.represent_mapping('tag:yaml.org,2002:map', data.items())

def excel_to_yml(excel_file_path, yml_file_path):

    df = pd.read_excel(excel_file_path, engine='openpyxl')

    data_dict = {
        "categories": []
    }

    # 按照类别进行分组处理
    for category_name, category_df in df.groupby('category_name'):
        category_dict = {
            "name": category_name,
            "subcategories": []
        }

        # 按照子类别进行分组处理
        for subcategory_name, subcategory_df in category_df.groupby('subcategory_name'):
            subcategory_dict = {
                "name": subcategory_name,
                "items": []
            }

            for _, row in subcategory_df.iterrows():
                item_dict = {
                    "name": row['item_name'],
                    "description": row['description'],
                    "homepage_url": row['homepage_url'] if not pd.isna(row['homepage_url']) else None,
                    "project": row['project'],
                    "repo_url": row['repo_url'],
                    "extra": {
                        "organization": row['organization']
                    }
                }
                if 'ohpm_url' in row and not pd.isna(row['ohpm_url']):
                    item_dict["extra"]["ohpm_url"] = row['ohpm_url']

                subcategory_dict["items"].append(item_dict)

            subcategory_dict["items"] = sorted(subcategory_dict["items"], key=lambda x: x["name"])
            category_dict["subcategories"].append(subcategory_dict)

        data_dict["categories"].append(category_dict)

    OrderedDumper.add_representer(dict, represent_dict_order)

    # 将字典转换为YAML格式并保存到文件
    with open(yml_file_path, 'w', encoding='utf-8') as yml_file:
        yaml.dump(data_dict, yml_file, Dumper=OrderedDumper, allow_unicode=True)


if __name__ == "__main__":
    excel_file_path = "/home/guoqiangqi/landscape2/landscape_debug/projects_list.xlsx"
    yml_file_path = "/home/guoqiangqi/landscape2/landscape_debug/projects_list.yml"
    excel_to_yml(excel_file_path, yml_file_path)