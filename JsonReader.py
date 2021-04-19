import json
import time

with open('.\\data/Result.json','r') as jr:
    output = json.load(jr)
    jr.close()

print(output)
temps = input('\n')