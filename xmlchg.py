

import xmltodict


def print_hi():
    path_to_presentation = './zip_file/ppt/slides/slide19.xml'
    with open(path_to_presentation, "rb") as file:
        doc = xmltodict.parse(file, process_namespaces=True)
        print(doc.keys())


if __name__ == '__main__':
    print_hi()

