"""BrickBridge"""
def brick_bridge(bridge, brick_l, brick_s):
    """a"""
    ans1 = bridge//5
    if ans1 > brick_l:
        ans1 = brick_l
    if brick_s - ans1*5 > bridge:
        print("-1")
    else:
        print(brick_s-(ans1*5))
brick_bridge(int(input()), int(input()), int(input()))

Hello
