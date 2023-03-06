import glob
import numpy as np
import matplotlib
import matplotlib.pylab as plt
from sklearn import manifold
import chainer
import chainer.functions as F
import chainer.links as L
from PIL import Image
from tqdm import tqdm
vgg = L.VGG16Layers()
vgg._children
paths = []
features = []
for path in tqdm(glob.glob('../../umemoto_image/*.jpg')):
    paths.append(path)
    img = Image.open(path)
    if img.mode != 'RGB':
        img = img.convert('RGB')
    feature = vgg.extract([img], layers=['pool5'])['pool5']
    feature = feature.data.reshape(-1)
    features.append(feature)
features = np.array(features)
def cos_sim_matrix(matrix):
    #print(matrix.shape)
    d = matrix @ matrix.T
    norm = (matrix * matrix).sum(axis=1, keepdims=True) ** .5
    return d / norm / norm.T
cos_sims = cos_sim_matrix(features)
samples = np.random.randint(0, len(paths), 10)
#print(features)
#print(np.linalg.norm(features[0], ord=0))
for i in samples:
    sim_idxs = np.argsort(cos_sims[i])[::-1]
    sim_idxs = np.delete(sim_idxs, np.where(sim_idxs==i))
    sim_num = 18
    sim_idxs = sim_idxs[:sim_num]
    fig, axs = plt.subplots(ncols=sim_num+1, figsize=(15, sim_num))
    img = Image.open(paths[i])
    axs[0].imshow(img)
    axs[0].set_title('target')
    axs[0].axis('off')
    for j in range(sim_num):
        img = Image.open(paths[sim_idxs[j]])
        axs[j+1].imshow(img)
        axs[j+1].set_title(cos_sims[i, sim_idxs[j]])
        axs[j+1].axis('off')
    #plt.show()