package oger.entity;


import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.util.Date;

/**
 * @Auther: Oger
 * @Date: 2020-07-22
 * @Description:
 */
@Data
@ApiModel("学生实体类")
public class Student {

    public Student(Integer id, String name, Date birthday) {
        this.id = id;
        this.name = name;
        this.birthday = birthday;
    }

    @ApiModelProperty("ID")
    private Integer id;

    @ApiModelProperty("姓名")
    private String name;

    @ApiModelProperty("生日")
    private Date birthday;
}
