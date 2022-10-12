def read_rtf(file:str, headers=True)->list[str]:
    with open(file, 'r') as file:
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
    if not headers:
        store = text2[0][83:87]
        header1 = text2[0][:70]
        header2 = text2[1][:70]
        header3 = text2[2][:70]
        i = 0
        while i < len(text2):
            line = text2[i]
            if header1 in line or header2 in line or header3 in line:
                text2.pop(i)
                i-=1
            i+=1 
        return text2, store


    return text2

if __name__ == '__main__': 
    text, store = read_rtf("C:\\ZocDownload\\DSHWKC.rtf",headers=False)
    print(store)
    for line in text:
        print(line)
    print(store)