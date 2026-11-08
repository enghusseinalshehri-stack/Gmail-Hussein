import json, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from PIL import Image
matplotlib.rcParams["mathtext.fontset"] = "cm"
COL = "#F1F5F9"
F = {
 "f1": (r"$F_r=\left[\frac{12}{nk(k+1)}\sum_{i=1}^{k}R_i^2\right]-3n(k+1)$", 30),
 "f1s": (r"$F_r=\left[\frac{12}{nk(k+1)}\sum_{i=1}^{k}R_i^2\right]-3n(k+1)$", 21),
 "f2": (r"$F_r=\frac{n(k-1)\left[\frac{\sum_{i=1}^{k}R_i^2}{n}-C_F\right]}{\sum r_{ij}^2-C_F}$", 26),
 "f2s": (r"$F_r=\frac{n(k-1)\left[\frac{\sum_{i=1}^{k}R_i^2}{n}-C_F\right]}{\sum r_{ij}^2-C_F}$", 19),
 "f3": (r"$C_F=\frac{1}{4}nk(k+1)^2$", 26),
 "df": (r"$df=k-1$", 26),
 "dfs": (r"$df=k-1$", 20),
 "h0": (r"$H_0:\ \theta_B=\theta_{M1}=\theta_{M2}$", 24),
 "ha": (r"$H_A:$", 24),
 "c1": (r"$F_r=\left[\frac{12}{7(3)(3+1)}\right](19^2+11^2+12^2)-3(7)(3+1)$", 26),
 "c2": (r"$=\left(\frac{12}{84}\right)(361+121+144)-84$", 26),
 "c3": (r"$=(0.1429)(626)-84$", 26),
 "c4": (r"$=89.4286-84$", 26),
 "c5": (r"$F_r=5.429$", 30),
 "sFr": (r"$F_r$", 20), "sn": (r"$n$", 20), "sk": (r"$k$", 20),
 "sRi": (r"$R_i$", 20), "sCF": (r"$C_F$", 20), "srij": (r"$r_{ij}$", 20),
}
sizes = {}
for k,(t,fs) in F.items():
    fig = plt.figure(figsize=(0.1,0.1))
    fig.text(0,0,t,fontsize=fs,color=COL)
    p=f"build/f/{k}.png"
    fig.savefig(p,dpi=300,transparent=True,bbox_inches="tight",pad_inches=0.03)
    plt.close(fig)
    w,h = Image.open(p).size
    sizes[k]={"w":w/300,"h":h/300}
json.dump(sizes,open("build/formulas.json","w"))
print({k:(round(v["w"],2),round(v["h"],2)) for k,v in sizes.items()})
