# excel-simple-download module

Excel 시트 생성 및 Excel 관련 기능을 빠르게 구현하는 Excel module입니다.

---

## 기본 필수 구현

`ExcelFile<T>`를 상속 받아 `renderExcel(List<T> list)`을 구현할 수 있습니다. 또한 생성자를 통해 원하는 `WorkBook`구현을 선택할 수 있습니다. 만약 `List<T> data` 에 대한 검증이 필요하다면 `validate(List<T> data)` 를 재정의하면 됩니다.

`renderExcel(List<T> list)`재정의에서는 다음의 3가지를 사용할 수 있습니다.

`renderHeader(Sheet sheet, int rowIndex)` : 필드이름 삽입 시점 및 필드이름 위치 결정

`renderBody(Sheet sheet, int rowIndex, T data)` : 데이터 셀 삽입 시점 및 데이터 셀 삽입 위치 결정

`renderFormula(Sheet sheet, int criteriaRowIndex, List<T> data)` : 수식이 적용된 셀 삽입 시점  및 위치 결정

```java
public class MyExcelFile<T> extends ExcelFile<T> {

    protected MyExcelFile(List<T> data, Class<?> clazz) {
        super(data, clazz, new SXSSFWorkbook());
    }

    @Override
    protected void validate(List<T> data) {
        // do nothing
    }

    @Override
    protected void renderExcel(List<T> list) {
        Workbook workbook = this.getWorkbook();
        Sheet sheet = workbook.createSheet();

        renderHeader(sheet, 0); // 필드 명 삽입 시점과 위치 결정
        for (int i = 1; i < list.size(); i++) {
					renderBody(sheet, i, list.get(i)); // 데이터 셀 삽입 시점과 위치 결정
        }
    }
}
```

---

## 간단하게 사용하기

필수 구현이 완료되었다면 `@ExcelColumn`을 사용하여 `WorkBook`을 렌더링 할 수 있습니다.

아래 예시를 참고 하세요.

```java
@Builder
public class ExcelDto {
    @ExcelColumn(headerName = "연번코드")
    private String indexCode;

    @ExcelColumn(headerName = "나이")
    private int age;

    @ExcelStyle(format = "yyyy-mm-dd")
    @ExcelColumn(headerName = "생일")
    private LocalDate birthDay;
}
```

```java
public class Application {
    public static void main(String[] args) {
        String indexCodeFormat = "%s-%s";
        LocalDate now = LocalDate.now();

        List<ExcelDto> targetDto = IntStream.rangeClosed(0, 10)
                .mapToObj(operand -> new ExcelDto.ExcelDtoBuilder()
                        .indexCode(String.format(indexCodeFormat, operand, operand))
                        .age(operand)
                        .birthDay(now.minusDays(operand))
                        .build()
                ).collect(Collectors.toList());

        MyExcelFile<ExcelDto> myExcelFile = new MyExcelFile<ExcelDto>(targetDto,ExcelDto.class);
        Workbook workbook = myExcelFile.getWorkbook();
    }
}
```

![Untitled](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/b987a1da-7927-4f45-82d2-88cc81615383/Untitled.png)

위와 같이 `@ExcelColumn`을 사용하여 간단히 Excel의 필드이름을 지정할 수 있습니다.

---

## 커스텀 옵션 제공

---

### `@ExcelFormula`

만약 아래와 같이 수식을 삽입해야한다면 `@ExcelFormula` 를 사용하면 됩니다. `@ExcelFormula(foramulaClass = class)`에 지정된 클래스를 찾아 실행 시킵니다. `formulaClass` 옵션을 사용하려면 `Formula interface`를 상속받아 사용하세요.

![Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/39b23485-c21a-4061-94cf-184d27843fce/Untitled.png)

```java
@Builder
**@ExcelFormula(formulaClass = AverageFormula.class)**
public class ExcelDto {
    @ExcelColumn(headerName = "연번코드")
    private String indexCode;

    @ExcelColumn(headerName = "나이")
    private int age;

    @ExcelStyle(format = "yyyy-mm-dd")
    @ExcelColumn(headerName = "생일")
    private LocalDate birthDay;
}
```

### 필수 구현 및 `renderFormula(Sheet sheet, int criteriaRowIndex, List<T> data)` 실행

```java
public class AverageFormula implements Formula {

    @Override
    public void validate(SheetHelper sheet, int criteriaRowIndex, List<?> data) {}

    @Override
    public void render(SheetHelper sheet, int criteriaRowIndex, List<?> data) {
        String formula = "AVERAGE(%s:%s)";
        String start = sheet.getCellAddressByFieldName(criteriaRowIndex, "age");
        String end = sheet.getCellAddressByFieldName(data.size()-1, "age");

        Cell targetCell = sheet.getCellByFieldName(data.size(), "age");

        sheet.setFormulaCell(targetCell,String.format(formula, start, end));
    }
}
```

```java
public class MyExcelFile<T> extends ExcelFile<T> {

    protected MyExcelFile(List<T> data, Class<?> clazz) {
        super(data, clazz, new SXSSFWorkbook());
    }

    @Override
    protected void validate(List<T> data) {
        // do nothing
    }

    @Override
    protected void renderExcel(List<T> list) {
        Workbook workbook = this.getWorkbook();
        Sheet sheet = workbook.createSheet();

        renderHeader(sheet, 0);
        for (int i = 1; i < list.size(); i++) {
            renderBody(sheet, i, list.get(i));
        }
        **renderFormula(sheet, 1, list); // 수식 사용을 위하여 시점과 위치 지정**
    }
}

```

---

### `@ExcelStyle`

만약 아래와 같이 스타일 설정이 필요하다면 `@ExcelStyle`을 사용하면 됩니다. `@ExcelStyle`에는 총 3가지 옵션이 존재하며 어노테이션의 위치에 따라 Sheet에 대한 옵션과 Column에 대한 옵션으로 나누어 집니다. 

![Untitled](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/174d3f44-778f-4ffa-b15e-2f69de79d10e/Untitled.png)

```java
@Builder
@ExcelFormula(formulaClass = AverageFormula.class)
**@ExcelStyle(headerStyleClass = DefaultHeaderStyle.class, bodyStyleClass = DefaultBodyStyle.class)**
public class ExcelDto {
    @ExcelColumn(headerName = "연번코드")
    private String indexCode;

    **@ExcelStyle(headerStyleClass = BackgroundBlue.class, bodyStyleClass = DefaultHeaderStyle.class)**
    @ExcelColumn(headerName = "나이")
    private int age;

    **@ExcelStyle(format = "yyyy-mm-dd")**
    @ExcelColumn(headerName = "생일")
    private LocalDate birthDay;
}
```

```java
// 클래스 사용
@ExcelStyle(
headerStyleClass = , // sheet에 삽입된 모든 필드이름에 적용되는 스타일
bodyStyleClass = , //  sheet에 삽입된 모든 데이터 셀에 적용되는 스타일
format = ) //sheet에 삽입된 모든 데이터 셀에 적용되는 데이터 포멧
public class ExcelDto{
		
		@ExcelStyle(headerStyleClass = , // 현재 column에 삽입된 필드이름에만 적용되는 스타일
								bodyStyleClass = , // 현재 column에 삽입된 데이터 셀에만 적용되는 스타일
								format = ) // 현재 column에 삽입된 데이터 셀에만 적용되는 데이터 포멧
    int val;
}
```

### 필수 구현 및 기본 설정 사용 방법

커스텀 `style`을 사용하려면 `Style interface`를 재정의하면 됩니다.

**사용방법**

```java
public class CustomStyle implements Style {

    @Override
    public boolean enableDefaultStyle() { // class에 사용된 Style 정보를 사용할지 여부 기본 설정 false
        return Style.super.enableDefaultStyle(); // true or false
    }

    @Override
    public void configure(CellStyle cellStyle) {
				// cellStyle 지정
    }
}
```

**실제 `Style`재정의**

```java
public class DefaultHeaderStyle implements Style {
    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setBorderLeft(BorderStyle.DOUBLE);
        cellStyle.setBorderRight(BorderStyle.DOUBLE);
        cellStyle.setBorderTop(BorderStyle.DOUBLE);
        cellStyle.setBorderBottom(BorderStyle.DOUBLE);
    }
}
```

```java
public class DefaultBodyStyle implements Style
    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
    }
}
```

```java
public class BackgroundBlue implements Style {
    @Override
    public boolean enableDefaultStyle() { // 만약 기본 설정에서 추가 하고 싶다면 true로 전환
        return true;
    }

    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

```
Excel 시트 생성 및 Excel 관련 기능을 빠르게 구현하는 Excel module입니다.

---

## 기본 필수 구현

`ExcelFile<T>`를 상속 받아 `renderExcel(List<T> list)`을 구현할 수 있습니다. 또한 생성자를 통해 원하는 `WorkBook`구현을 선택할 수 있습니다. 만약 `List<T> data` 에 대한 검증이 필요하다면 `validate(List<T> data)` 를 재정의하면 됩니다.

`renderExcel(List<T> list)`재정의에서는 다음의 3가지를 사용할 수 있습니다.

`renderHeader(Sheet sheet, int rowIndex)` : 필드이름 삽입 시점 및 필드이름 위치 결정

`renderBody(Sheet sheet, int rowIndex, T data)` : 데이터 셀 삽입 시점 및 데이터 셀 삽입 위치 결정

`renderFormula(Sheet sheet, int criteriaRowIndex, List<T> data)` : 수식이 적용된 셀 삽입 시점  및 위치 결정

```java
public class MyExcelFile<T> extends ExcelFile<T> {

    protected MyExcelFile(List<T> data, Class<?> clazz) {
        super(data, clazz, new SXSSFWorkbook());
    }

    @Override
    protected void validate(List<T> data) {
        // do nothing
    }

    @Override
    protected void renderExcel(List<T> list) {
        Workbook workbook = this.getWorkbook();
        Sheet sheet = workbook.createSheet();

        renderHeader(sheet, 0); // 필드 명 삽입 시점과 위치 결정
        for (int i = 1; i < list.size(); i++) {
					renderBody(sheet, i, list.get(i)); // 데이터 셀 삽입 시점과 위치 결정
        }
    }
}
```

---

## 간단하게 사용하기

필수 구현이 완료되었다면 `@ExcelColumn`을 사용하여 `WorkBook`을 렌더링 할 수 있습니다.

아래 예시를 참고 하세요.

```java
@Builder
public class ExcelDto {
    @ExcelColumn(headerName = "연번코드")
    private String indexCode;

    @ExcelColumn(headerName = "나이")
    private int age;

    @ExcelStyle(format = "yyyy-mm-dd")
    @ExcelColumn(headerName = "생일")
    private LocalDate birthDay;
}
```

```java
public class Application {
    public static void main(String[] args) {
        String indexCodeFormat = "%s-%s";
        LocalDate now = LocalDate.now();

        List<ExcelDto> targetDto = IntStream.rangeClosed(0, 10)
                .mapToObj(operand -> new ExcelDto.ExcelDtoBuilder()
                        .indexCode(String.format(indexCodeFormat, operand, operand))
                        .age(operand)
                        .birthDay(now.minusDays(operand))
                        .build()
                ).collect(Collectors.toList());

        MyExcelFile<ExcelDto> myExcelFile = new MyExcelFile<ExcelDto>(targetDto,ExcelDto.class);
        Workbook workbook = myExcelFile.getWorkbook();
    }
}
```

![Untitled](https://user-images.githubusercontent.com/66084125/220286409-12ce5d8e-cf00-4dd6-b489-f8a59249c74c.png)

위와 같이 `@ExcelColumn`을 사용하여 간단히 Excel의 필드이름을 지정할 수 있습니다.

---

## 커스텀 옵션 제공

---

### `@ExcelFormula`

만약 아래와 같이 수식을 삽입해야한다면 `@ExcelFormula` 를 사용하면 됩니다. `@ExcelFormula(foramulaClass = class)`에 지정된 클래스를 찾아 실행 시킵니다. `formulaClass` 옵션을 사용하려면 `Formula interface`를 상속받아 사용하세요.

![Untitled 1](https://user-images.githubusercontent.com/66084125/220286466-754023fa-a356-4443-b1ba-8f2255259f0d.png)

```java
@Builder
**@ExcelFormula(formulaClass = AverageFormula.class)**
public class ExcelDto {
    @ExcelColumn(headerName = "연번코드")
    private String indexCode;

    @ExcelColumn(headerName = "나이")
    private int age;

    @ExcelStyle(format = "yyyy-mm-dd")
    @ExcelColumn(headerName = "생일")
    private LocalDate birthDay;
}
```

### 필수 구현 및 `renderFormula(Sheet sheet, int criteriaRowIndex, List<T> data)` 실행

```java
public class AverageFormula implements Formula {

    @Override
    public void validate(SheetHelper sheet, int criteriaRowIndex, List<?> data) {}

    @Override
    public void render(SheetHelper sheet, int criteriaRowIndex, List<?> data) {
        String formula = "AVERAGE(%s:%s)";
        String start = sheet.getCellAddressByFieldName(criteriaRowIndex, "age");
        String end = sheet.getCellAddressByFieldName(data.size()-1, "age");

        Cell targetCell = sheet.getCellByFieldName(data.size(), "age");

        sheet.setFormulaCell(targetCell,String.format(formula, start, end));
    }
}
```

```java
public class MyExcelFile<T> extends ExcelFile<T> {

    protected MyExcelFile(List<T> data, Class<?> clazz) {
        super(data, clazz, new SXSSFWorkbook());
    }

    @Override
    protected void validate(List<T> data) {
        // do nothing
    }

    @Override
    protected void renderExcel(List<T> list) {
        Workbook workbook = this.getWorkbook();
        Sheet sheet = workbook.createSheet();

        renderHeader(sheet, 0);
        for (int i = 1; i < list.size(); i++) {
            renderBody(sheet, i, list.get(i));
        }
        **renderFormula(sheet, 1, list); // 수식 사용을 위하여 시점과 위치 지정**
    }
}

```

---

### `@ExcelStyle`

만약 아래와 같이 스타일 설정이 필요하다면 `@ExcelStyle`을 사용하면 됩니다. `@ExcelStyle`에는 총 3가지 옵션이 존재하며 어노테이션의 위치에 따라 Sheet에 대한 옵션과 Column에 대한 옵션으로 나누어 집니다. 

![Untitled 2](https://user-images.githubusercontent.com/66084125/220286486-b9bef40c-bd9c-49ea-964d-1e8be02ec9c8.png)

```java
@Builder
@ExcelFormula(formulaClass = AverageFormula.class)
**@ExcelStyle(headerStyleClass = DefaultHeaderStyle.class, bodyStyleClass = DefaultBodyStyle.class)**
public class ExcelDto {
    @ExcelColumn(headerName = "연번코드")
    private String indexCode;

    **@ExcelStyle(headerStyleClass = BackgroundBlue.class, bodyStyleClass = DefaultHeaderStyle.class)**
    @ExcelColumn(headerName = "나이")
    private int age;

    **@ExcelStyle(format = "yyyy-mm-dd")**
    @ExcelColumn(headerName = "생일")
    private LocalDate birthDay;
}
```

```java
// 클래스 사용
@ExcelStyle(
headerStyleClass = , // sheet에 삽입된 모든 필드이름에 적용되는 스타일
bodyStyleClass = , //  sheet에 삽입된 모든 데이터 셀에 적용되는 스타일
format = ) //sheet에 삽입된 모든 데이터 셀에 적용되는 데이터 포멧
public class ExcelDto{
		
		@ExcelStyle(headerStyleClass = , // 현재 column에 삽입된 필드이름에만 적용되는 스타일
								bodyStyleClass = , // 현재 column에 삽입된 데이터 셀에만 적용되는 스타일
								format = ) // 현재 column에 삽입된 데이터 셀에만 적용되는 데이터 포멧
    int val;
}
```

### 필수 구현 및 기본 설정 사용 방법

커스텀 `style`을 사용하려면 `Style interface`를 재정의하면 됩니다.

**사용방법**

```java
public class CustomStyle implements Style {

    @Override
    public boolean enableDefaultStyle() { // class에 사용된 Style 정보를 사용할지 여부 기본 설정 false
        return Style.super.enableDefaultStyle(); // true or false
    }

    @Override
    public void configure(CellStyle cellStyle) {
				// cellStyle 지정
    }
}
```

**실제 `Style`재정의**

```java
public class DefaultHeaderStyle implements Style {
    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setBorderLeft(BorderStyle.DOUBLE);
        cellStyle.setBorderRight(BorderStyle.DOUBLE);
        cellStyle.setBorderTop(BorderStyle.DOUBLE);
        cellStyle.setBorderBottom(BorderStyle.DOUBLE);
    }
}
```

```java
public class DefaultBodyStyle implements Style
    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
    }
}
```

```java
public class BackgroundBlue implements Style {
    @Override
    public boolean enableDefaultStyle() { // 만약 기본 설정에서 추가 하고 싶다면 true로 전환
        return true;
    }

    @Override
    public void configure(CellStyle cellStyle) {
        cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

```
