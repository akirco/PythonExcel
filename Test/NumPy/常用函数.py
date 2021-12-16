import numpy as np

i2 = np.eye(2)
print(i2)

np.savetxt('eye.txt', i2)

# c, v = np.loadtxt('data.csv', delimiter=',', usecols=(6, 7), unpack=True)
# vwap = np.average(c, weights=v)
# print("VWAP =", vwap)
# VWAP = 350.589549353
# print(np.mean(c))
#
# t = np.arange(10)
# print('twap:', np.average(c, weights=t))
