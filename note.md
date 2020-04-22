# Bond Trading System

## 0 Three Key Factors

- Expression skill
- Business knowledge
- Expertise

## 1 Time Value of Money

- Simple Interest
- Compound Interest(**capitallsed**-same interest rate)

**Effective Rate**

- $ER=(1+NR/t)^t-1$

**Feature Value**

- $FV=PV\times (1+R/t)^{N\times t}$

- $\text{Discount factor} = (1+R/t)^{N\times t}$

Risk: interest rate is changing, reinvest risk

Use of PV and FV:

- Pricing
- Investment decision

**The pricing formula**
$$
\begin{equation}\begin{split}
\text{Bond Price} &= \frac{C/t}{(1+R/t)}+\frac{C/t}{(1+R/t)^2}+\cdots+\frac{\text{Principal}+C/t}{(1+R/t)^{n\times t}}\\
&= C/t\times \sum_{i=1}^{n}\frac{1}{(1+R/t)i}+\frac{\text{Principal}}{(1+R/t)^{n\times t}}\\
&= C/t\times \frac{D_1-D_{n\times t+1}}{1-D_1}+\text{Principal}\times D_{n\times t}\\
& where\, D_i=\frac{1}{(1+R/t)^i}
\end{split}\end{equation}
$$


**PV of Perpetuity**
$$
\begin{equation}\begin{split}
\text{Bond Price} &= C/t\times \frac{D_1(1-D_{n\times t})}{1-D_1}+\text{Principal}\times D_{n\times t}\\
&= \frac{C}{R}\text{  when n approaches infinity}
\end{split}\end{equation}
$$


## 2 Bond Basis

### 2.1 Valuation Formula

**Pricing off a Coupon Date**
$$
\begin{equation}\begin{split}
PV &= \frac{C/t}{(1+R/t)^a}+\frac{C/t}{(1+R/t)^{a+1}}+\cdots+\frac{\text{Principal}+C/t}{(1+R/t)^{a+m}}\\
&= C/t\times \frac{D_a-D_{a+m+1}}{1-D_1}+\text{Principal}\times D_{a+m}\\
m&= \text{complete coupond periods}\\
a&=\frac{\text{Number of days to next coupon}}{\text{Number of days in current coupond period}}
\end{split}\end{equation}
$$

**Accrued Interest**

- Dirty price = Settlement price = PV

- Clean price = Dirty price - Accrued interest

- Accrued interest = $C/t\times \text{Fractional coupon period}$

- Fractional coupon period = $\frac{\text{Numer of days since last coupon}}{\text{Number of days in current counpon period}}$

- Principal Conventions

  - Actual/Actual
  - Actual/365
  - 30/360

**Yield To Maturity**

The same formula as PV, but `R` is the target now.

**Macaulay Duration**
$$
\begin{equation}\begin{split}
D = \frac{1}{PV}\times [(a)\times\frac{C/t}{(1+R/t)^a}+(a+1)\times\frac{C/t}{(1+R/t)^{a+1}}+\cdots+(a+m)\times\frac{\text{Principal}+C/t}{(1+R/t)^{a+m}}]
\end{split}\end{equation}
$$
**Modified Duration**
$$
\begin{equation}\begin{split}
\text{Modified Duration} = \frac{\text{Macaulay Duration(in years)}}{1 + R/t}
\end{split}\end{equation}
$$
**Basis Point Value**
$$
\begin{equation}\begin{split}
BPV = \frac{\text{Modified Duration}(\%)}{100}\times\frac{\text{Dirty Price}}{100}
\end{split}\end{equation}
$$
