package oger.entity;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.util.List;

/**
 * @Auther: Oger
 * @Date: 2020-09-27
 * @Description:
 */
@Data
@ApiModel("学生信息实体类")
public class StudentInfo {

    @ApiModelProperty("班级")
    private String classes;

    @ApiModelProperty("成绩")
    private List<Student> scores;

    @ApiModelProperty("语文总成绩")
    private int totalChineseScore;

    @ApiModelProperty("数学总成绩")
    private int totalMathScore;

    @ApiModelProperty("总成绩")
    private int totalScore;
}
