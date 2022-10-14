def read_rtf(rtf_file:str)->list[str]:
    with open(rtf_file, 'r') as file:
        text = file.readlines()
    for i in range(0,5):
        text.pop(0)
    text2 = []
    for line in text:
        if len(line) != 5:
            if "\\page" in line:
                text2.append("")
            else:
                text2.append(line[66:-2])
    del(text)
    return text2

