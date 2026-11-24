# ---
# jupyter:
#   jupytext:
#     formats: ipynb,jl:percent
#     text_representation:
#       extension: .jl
#       format_name: percent
#       format_version: '1.3'
#       jupytext_version: 1.19.7
#   kernelspec:
#     display_name: Julia
#     language: julia
#     name: julia-1.10
# ---

# %% [markdown]
# ## Julia example
#
# Text.

# %%
julia_text = """start
# %% not a boundary
end"""
length(julia_text)

# %%
sum(1:5)
