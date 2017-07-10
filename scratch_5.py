import pandas as pd
import numpy as np

df_ = pd.DataFrame([['yes','yes','no','yes'],['yes','no','yes','yes'],['no','yes','yes','no'],['yes','yes','no','yes']],columns=["Address","City/Zip","Country","Phone"])
print df_