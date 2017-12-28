package org.springframework.view.extension.config;


/**
 * Annotation to mark a field of <ExcelBean> as a column in resultant excel file.
 * 
 * @author nitish
 *
 */
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.METHOD)
public @interface DynamicExcelHeader {
    
    /**
     * @return - The columnNumber.
     */
    int columnNumber() default 0;
    int rowNumber() default 0;
    /**
     * @return - The columnNumber.
     */
    String colorFormatter() default "";
}
