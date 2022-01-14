package com.mmg;

import lombok.Data;

import java.util.Date;
import java.util.List;

/**
 * @Auther: fan
 * @Date: 2021/11/29
 * @Description:
 */
@Data
public class ExecutionPdfVO {

    private String cropNameCn;
    private String cropId;
    private List<DataInfo> info;

    @Data
    public static class DataInfo {
        private String formulaName;
        private String lureNameCn;
        private String lureId;
        private List<LureInfo> lureInfoList;
    }

    @Data
    public static class LureInfo {
        private String executionResult;
        private Date startTime;
        private Date endTime;
    }
}
