# Scripts
fprinciple_brief.ipynb: A simple code to verify frequency principle for 1d functions.

fprinciple.ipynb: A complete code to verify frequency principle for 1d functions.

f-principle_nd.ipynb: A code to verify frequency principle for n-dimensonal functions.

example.ipynb: Answers to tasks in 1d scripts.

# Related papers

[1] Zhi-Qin John Xu* , Yaoyu Zhang, and Yanyang Xiao, Training behavior of deep neural network in frequency domain, arXiv preprint: 1807.01251, (2018), 26th International Conference on Neural Information Processing (ICONIP 2019). [pdf](https://ins.sjtu.edu.cn/people/xuzhiqin/pub/training_behavior_ICONIP2019_XZX.pdf) and [web](https://link.springer.com/chapter/10.1007/978-3-030-36708-4_22)

[2] Zhi-Qin John Xu* , Yaoyu Zhang, Tao Luo, Yanyang Xiao, Zheng Ma, Frequency Principle: Fourier Analysis Sheds Light on Deep Neural Networks, arXiv preprint: 1901.06523, Communications in Computational Physics (CiCP). [pdf](https://ins.sjtu.edu.cn/people/xuzhiqin/pub/shedlightCiCP.pdf) and in [web](https://www.global-sci.org/intro/article_detail/cicp/18395.html), some code is in [github](https://github.com/xuzhiqin1990/F-Principle) 

[3] Zhi-Qin John Xu*, Yaoyu Zhang, Tao Luo, Overview frequency principle/spectral bias in deep learning. arxiv 2201.07395 (2022) . [pdf](https://ins.sjtu.edu.cn/people/xuzhiqin/pub/fpoverview2201.07395.pdf), and in [arxiv](https://arxiv.org/abs/2201.07395).

For more details, refer to [Zhi-Qin John Xu's homepage](https://ins.sjtu.edu.cn/people/xuzhiqin/pub.html)

A bilibili course that is helpful for learning python: [Bilibili course](https://www.bilibili.com/video/BV16H4y1Q7tj/?p=1&vd_source=9e3c7a35167d2d11f2549c94242850e1)

# Frequency Principle (1D)

DNNs often fit target functions from low to high frequencies.  

The first figure shows the evolution of the function in spatial domain, the red line is the target function, and the blue line is the DNN output. *Ordinate vs. Abscissa : y vs. x*. 

The second figure shows the evolution of the function in Fourier domain, the red line is the FFT of the target function, and the blue line is the FFT of DNN output. *Ordinate vs. Abscissa: amplitude vs. frequency*. 

<!-- ![Title](https://ins.sjtu.edu.cn/people/xuzhiqin/index.html) -->
![value](./pic/value.gif)![freq](./pic/freq.gif)


### Core objective

Explore the generality of the frequency principle. Is it a universal law? What design should we do in the experiment?

### Report requirements

The relationship between the frequency principle and a series of hyperparameters. Choose one or two hyperparameters for experimental research in the report. Specific hyperparameter options can be found in the task section.

## Contact information

Zhi-Qin John Xu (许志钦): xuzhiqin@sjtu.edu.cn



