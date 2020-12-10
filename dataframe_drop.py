import pandas as pd
import os

df = pd.DataFrame(pd.np.arange(12).reshape(3, 4),
                  columns=['A', 'B', 'C', 'D'])

print(df)
df = df.drop([0, 1])
print(df)