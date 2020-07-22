package oger.entity;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

/**
 * @Auther: Oger
 * @Date: 2020-07-22
 * @Description:
 */
@Data
@ApiModel("教师实体类")
public class Teacher {

    public Teacher(Integer id, String name, String subject) {
        this.id = id;
        this.name = name;
        this.subject = subject;
    }

    @ApiModelProperty("ID")
    private Integer id;

    @ApiModelProperty("姓名")
    private String name;

    @ApiModelProperty("科目")
    private String subject;
}
