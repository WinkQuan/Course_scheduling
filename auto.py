import subprocess as sp

f = open("./result.txt", mode="w")

for i in range(50, 501, 50):
    cmd = r"python .\classtable_v2.py " + str(i)
    res = sp.getoutput(cmd)
    f.write(str(i)+"\n")
    f.write(res)
    f.write("\n\n")

f.close()