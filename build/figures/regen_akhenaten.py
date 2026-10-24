# Re-render the book's Akhenaten-eclipse map for the lecture, without any reference to the novel.
import sys, os
REPO = "/home/user/eclipse"
sys.path[:0] = [REPO, os.path.join(REPO, "book")]
import common, ch2_novel as N
OUT = sys.argv[1]
N.header_orig = N.header
def header(fig, title, sub):
    N.header_orig(fig, title, "كسوف كلي مرّ فوق أخت-آتون عاصمة إخناتون بعد الظهر — بالحساب الحديث")
N.header = header
N.wrap_orig = N.wrap
N.wrap = lambda s, n: N.wrap_orig(s.replace("أما نسبة الكسوف إلى السنة السادسة عشرة من حكم إخناتون فاجتهاد من الرواية لا يثبته نص مصري مؤرّخ.",
                                            "أما ربطه بسنة بعينها من حكم إخناتون فاجتهاد لا يثبته نص مصري مؤرّخ."), n)
def save(fig, chapter, slug, caption, kind=""):
    fig.savefig(OUT, dpi=common.DPI, facecolor=fig.get_facecolor()); common.plt.close(fig); print("saved", OUT)
N.save = save
D = N.compute()
N.fig_akhenaten(D["akh"])
