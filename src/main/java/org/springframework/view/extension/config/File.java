package org.springframework.view.extension.config;

/**
 * Annotation to mark a class of type <ReportBean> which corresponds to resultant report file i.e. csv or excel.
 * It holds meta information like template to be used for end result, title of the excel/csv report file, number of the row from which the data to be rendered, etc
 * 
 * @author nitish
 *
 */
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface File {

	/**
     * @return - The title.
     */
    String title() default "";
    boolean appendDate() default false;
    boolean appendExtension() default false;
    String template() default "";
    int dataRowStart() default 0;
    String sheetName() default "Sheet1";
    
    
    
}

