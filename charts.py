"""
Charts
waiky.jung@gmail.com
This program generates some random charts.
"""
import numpy as np
from matplotlib import pyplot as plt

#Subplot 1
sample1 = np.random.normal(loc=65, scale=4, size=300)
sample2 = np.random.normal(loc=75, scale=4, size=300)
ax1 = plt.subplot(2, 1, 1)
ax1.hist(sample1, bins=30, color='green', alpha=0.5)
ax1.hist(sample2, bins=30, color='orange', alpha=0.5)
xticks = range(55, 86, 5)
yticks = range(0, 40, 5)
ax1.set_xticks(xticks)
ax1.set_yticks(yticks)
ax1.set_xlabel('Samples', fontsize=9, loc='center', color='black')
ax1.set_ylabel('Count', fontsize=9, loc='center', color='black')
ax1.set_title('Random Samples', fontsize=11)

#Subplot 2
ax2 = plt.subplot(2, 2, 3)
x = [x for x in range(1, 10, 1)]
y = [2 * y for y in x]
ax2.plot(range(len(x)), x, color='red', marker='o')
ax2.plot(range(len(y)), y, color='blue', marker='x')
xticks = range(9)
yticks = range(0,20,2)
ax2.set_xticks(xticks)
ax2.set_yticks(yticks)
ax2.set_xlabel('A Set of Numbers', fontsize=9, loc='center', color='black')
ax2.set_ylabel('Another Set of Numbers', fontsize=9, loc='center', color='black')
ax2.legend(['A Line', 'Another Line'], loc=0, fontsize=8)
ax2.set_title('A Couple of Lines', fontsize=11)

#Subplot 3
ax3 = plt.subplot(2, 2, 4)
names = ['Amount I Will Eat',
         'Amount My Wife Will Eat',
         'Amount I Will Eat (2nd Round)',
         'Rest of the Pizza We\'ll Throw Away']
values = [10, 3, 5, 82]
colors = ['blue',
          'red',
          'yellow',
          'orange']
ax3.pie(values,
        wedgeprops={'linewidth': 1, 'edgecolor': 'white'},
        colors=colors,
        autopct='%0.0f%%',
        pctdistance=1.2,
        textprops={'fontsize': 7},
        startangle=90,
        counterclock=False)
ax3.legend(names,
           loc='lower center',
           ncol=2,
           fontsize=6,
           bbox_to_anchor=(0.5, -0.20))
ax3.set_title('Pizza Consumption', fontsize=11)

#Display All Charts
plt.subplots_adjust(left=0.1,
                    bottom=0.1,
                    right=0.9,
                    top=0.9,
                    wspace=0.4,
                    hspace=0.4)
plt.show()
#plt.close('all')
