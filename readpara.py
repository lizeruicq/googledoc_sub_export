



def readpara(paraname):
    paralist = []
    para = open(paraname, 'r')
    contents = para.readlines()

    for line in contents:
        line =line.split(':')[-1]
        paralist.append(line.strip('\n'))
    return paralist

print("参数："+str(readpara('参数.txt')))
