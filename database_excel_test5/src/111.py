import pandas as pd
import numpy as np

df = pd.DataFrame(np.random.randint(0, 150, size=(5, 3)), columns=['a', 'b', 'c'])
print(df)

df2 = pd.concat((df, pd.DataFrame([])), keys=['A', ''], axis=1)
print(df2.columns)