package uk.co.terminological.tabular;

import java.io.ByteArrayOutputStream;
import java.io.PrintWriter;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import uk.co.terminological.datatypes.FluentSet;
import uk.co.terminological.mappers.ProxyMapWrapper;

public class JavaGenerator {

	/**
	 * Generated an interface spec based on a 
	 * @param className
	 * @param packageName
	 * @param extendsClass
	 * @param attrTypes
	 * @return
	 */
	public static String createInterface(
			String className, String packageName, List<Class<?>> extendsClass,
			Map<String,Class<?>> attrTypes) {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		PrintWriter out = new PrintWriter(baos);
		
		//pre-amble
		out.println("package "+packageName+";");
		out.println();
		for (Class<?> importClass: FluentSet.create(extendsClass, attrTypes.values())) {
			if (!importClass.getPackage().getName().equals("java.lang") &&
					!importClass.getPackage().getName().equals(packageName)
					) {
				out.println("import "+importClass.getCanonicalName()+";");
			}
		}
		out.println();

		//definition
		out.print("public interface "+className+" ");
		if (!extendsClass.isEmpty()) {
			out.print("extends ");
			out.print(extendsClass.stream().map(c -> c.getName()).collect(Collectors.joining(", ")));
			out.print(" ");
		}
		out.println("{");
		
		attrTypes.entrySet().stream()
			.forEach(at -> {
				out.print("\tpublic ");
				out.print(at.getValue().getSimpleName());
				out.print(" "+ProxyMapWrapper.methodName(at.getKey())+"();");
				out.println();
			});
		
		//close
		out.println("}");
		
		out.close();
		return baos.toString();
	}
	
		
	
	
	
}
