from gooey import Gooey, GooeyParser


@Gooey(
    program_name='图片合并pdf',
    menu=[{
        'name': '关于',
        'items': [{
            'type': 'AboutDialog',
            'menuTitle': '关于',
            'name': '图片合并PDF',
            'description': '将文件夹中图片生成pdf',
            'version': '1.0.0',
            'copyright': '2021',
            'website': 'https://github.com/wlzcool/ImageToPdf',
            'developer': 'https://juejin.cn/user/2815188501792951',
            'license': 'GNU'
        }, {
            'type': 'Link',
            'menuTitle': '程序下载地址',
            'url': 'https://github.com/wlzcool/ImageToPdf'
        }]
    }]
)
def init():
    parser = GooeyParser(description="本程序基于python编写，作者：Icebear")
    parser.add_argument('num', metavar='必填参数', help='请输入您要归类的直播(例：1号线，输入1回车):')
    args = parser.parse_args()
    num = args.num
    # # 1号线
    res1_list = (601, 602, 603, 604, 605, 606, 607, 608, 609, 610, 611, 612, 613, 614, 615, 616, 617, 618, 619, 620,
                 621, 622, 623, 624, 625, 626, 627, 628, 629, 630, 631, 632, 633, 634, 635, 636, 637, 638, 639, 640,
                 641, 642, 643, 644, 645, 646, 647, 648, 649, 650, 651, 652, 653, 654, 655, 656, 657, 658, 659, 724
                 )

    # # 2号线
    res2_list = (701, 702, 703, 704, 705, 706, 707, 708, 709, 710, 711, 712, 713, 714, 715, 716, 717, 718, 719, 720,
                 721, 722, 723, 724, 725, 726, 727, 728, 729, 730, 731, 732, 733, 734, 735, 736, 737, 738, 739, 740,
                 741, 742)

    # num = input('请输入您要归类的直播(例：1号线，输入1回车):')
    # print('已接受输入数据,正在处理......')
    if num == '1':
        print('已接受输入数据【1】,正在处理1号线数据......')
        return res1_list, num
    elif num == '2':
        print('已接受输入数据【2】,正在处理2号线数据......')
        return res1_list, num


if __name__ == '__main__':
    init()
