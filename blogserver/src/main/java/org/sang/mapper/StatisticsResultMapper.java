package org.sang.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.sang.bean.StatisticsResult;
import org.sang.bean.WorkRecord;

import java.util.List;
import java.util.Map;

@Mapper
public interface StatisticsResultMapper {
    List<StatisticsResult> getStatisticsResultList(@Param("name") String name);
    List<Map<String,Object>> getStatisticsResultListToExport(@Param("name") String name);

    int insertStatisticResult(StatisticsResult statisticsResult);

    List<Map<String,Object>> getAllResultList();
    List<Map<String,Object>> getAllResultListByMonth(@Param("name") String name);
    List<Map<String,Object>> getAllResultListByMonthToExport(@Param("name") String name);
    int  delectStatisticResult();
}
