## Steps


### Create conda env

```
conda create -n tqa Python=3.8
conda activate tqa
```

### Install Depencencies

```
pip install streamlit
pip install transformers
pip install torch torchvision torchaudio
pip install torch-scatter -f https://data.pyg.org/whl/torch-1.10.0+cpu.html

```

### To run the code

```
streamlit run app.py
```