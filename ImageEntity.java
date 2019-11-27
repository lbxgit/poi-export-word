package per.qiao.utils.hutool.poi;

import static per.qiao.utils.hutool.poi.ImageUtils.ImageType.PNG;

/**
 * Create by IntelliJ Idea 2018.2
 *
 * 图片实体对象
 *
 * @author: qyp
 * Date: 2019-10-26 21:52
 */
public class ImageEntity {

    /**
     * 图片宽度
     */
    private int width = 400;

    /**
     * 图片高度
     */
    private int height = 300;

    /**
     * 图片地址
     */
    private String url;

    /**
     * 图片类型
     * @see ImageUtils.ImageType
     */
    private ImageUtils.ImageType typeId = PNG;

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public ImageUtils.ImageType getTypeId() {
        return typeId;
    }

    public void setTypeId(ImageUtils.ImageType typeId) {
        this.typeId = typeId;
    }
}
