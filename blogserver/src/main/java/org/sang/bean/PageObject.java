package org.sang.bean;

import lombok.Data;

import java.io.Serializable;
import java.util.List;

/**
 * 借助此类封装业务层分页信息
 * 建议：所有用于封装数据的对象都实现Servializable接口（此接口是对象
 * 是否可以序列化的标识）
 * FAQ？
 * 1、何为序列化和反序列
 * 1）序列化：将对象转换为字节
 * 2）反序列化：将字节转化为对象
 * 2、序列化和反序列化应用场景
 * 1）将对象转换为字节存储到内存或文件
 * 2）将对象转为字节通过网络进行传输
 * 3）java中对象的序列化实现
 * 1）ObjectOutputStream(用于将对象序列化)
 * 2）ObjectInputStream(将字节反序列化为对象)
 * @param <T>
 */
@Data
public class PageObject<T> implements Serializable {
    //序列化和反序列化操作时的唯一标识，建议只要实现序列化接口，都要手动添加这个id
    private static final long serialVersionUID = -3130527491950235344L;
    /**总记录数*/
    private Integer rowCount;
    /**当前页记录*/
    private List<T> records;
    /**总页数*/
    private Integer pageCount;
    /**页面大小（每页最多显示多少条记录）*/
    private Integer pageSize;
    /**页码值*/
    private Integer pageCurrent;

    public PageObject() {}

    public PageObject(Integer rowCount, List<T> records, Integer pageSize, Integer pageCurrent) {
        this.rowCount = rowCount;
        this.records = records;
        this.pageSize = pageSize;
        this.pageCurrent = pageCurrent;
        this.pageCount = this.rowCount/this.pageSize;
        if (this.rowCount%this.pageSize!=0)
            this.pageCount++;
    }
}
