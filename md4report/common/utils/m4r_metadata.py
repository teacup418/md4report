import yaml



def get_metadata(file_path)->dict:
    '''
    读取文件开头的元数据
    '''
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
        
        # 提取 YAML Front Matter
        if content.startswith('---'):
            end = content.find('---', 3)
            if end != -1:
                yaml_content = content[3:end].strip()
                metadata = yaml.safe_load(yaml_content)
                return metadata
    
    return {}

# 示例用法
if __name__ == "__main__":
    metadata = get_metadata(r"C:\Users\Vincent\OneDrive\Documents\sync\obsidian\repo\数字图像处理.实验1.Python图像处理编程基础.md")
    print(metadata)