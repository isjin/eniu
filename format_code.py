f=open('STOCK.blk','r')
lines=f.readlines()
f.close()
f=open('stock.txt','w')
for line in lines:
    f.write(str(line)[1:])
f.close()