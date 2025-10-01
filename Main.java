import java.util.HashSet;
import java.util.Scanner;
import java.util.Set;

public class Main {

    public static void main(String[] args) {

        Scanner sc = new Scanner(System.in);
        String input = sc.nextLine();

        char op = ' ';
        if (input.contains("+")) op = '+';
        else if (input.contains("*")) op = '*';
        else if (input.contains("-")) op = '-';
        else {
            System.out.println("error");
            return;
        }

        String[] parts = input.split("\\" + op);
        if (parts.length != 2) {
            System.out.println("error");
            return;
        }

        Set<Integer> setA = parseSet(parts[0]);
        Set<Integer> setB = parseSet(parts[1]);

        Set<Integer> result = new HashSet<>(setA);
        switch (op) {
            case '+': result.addAll(setB); break; // Union
            case '*': result.retainAll(setB); break; // Intersection
            case '-': result.removeAll(setB); break; // Difference
        }

        System.out.println(result);
    }

    private static Set<Integer> parseSet(String s) {
        s = s.replaceAll("\\[|\\]", "").trim();
        Set<Integer> set = new HashSet<>();
        if (!s.isEmpty()) {
            String[] nums = s.split(",");
            for (String n : nums) {
                set.add(Integer.parseInt(n.trim()));
            }
        }
        return set;
    }

    }

