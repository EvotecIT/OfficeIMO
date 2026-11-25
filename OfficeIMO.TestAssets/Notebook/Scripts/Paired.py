# ---
# jupyter:
#   jupytext:
#     cell_metadata_filter: all
#     formats: ipynb,py:percent
#     text_representation:
#       extension: .py
#       format_name: percent
#       format_version: '1.3'
#       jupytext_version: 1.19.7
#   kernelspec:
#     display_name: Python 3
#     language: python
#     name: python3
# ---

# %% [markdown]
# ## Result
#
# A notebook with saved outputs.

# %% attributes={"classes": [], "id": "", "n": "10"} trusted=true
sample = """line 1
# %% [markdown] innocent text in a string
line 3"""
print(sample)

# %%
# %%time
print('second cell')

# %% [raw]
# Appendix notes
