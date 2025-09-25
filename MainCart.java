package p12_1;

import java.io.ObjectInputStream.GetField;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;

public class Main {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		ArrayList<Item> itemList = new ArrayList<Item>();
		String conshop;
		
		do {
		System.out.printf("Enter he name of the item: ");
		Scanner sc = new Scanner(System.in);
		String name = sc.nextLine();
		System.out.printf("Enter the unit price: ");
		double price = sc.nextDouble();
		System.out.printf("Enter the quantity: ");
		int quantity = sc.nextInt();
		sc.nextLine();
		System.out.printf("Continue shopping (y/n)? ");
		conshop = sc.nextLine();
		Item item = new Item(name, price, quantity);
		itemList.add(item);
		}while(conshop.equals("y"));
		
		System.out.println("Final Shopping Cart totals");
		double total=0;
		for(Item i:itemList) {
			System.out.println(i.getName() + "\t" + i.getPrice() + "$\t" + i.getQuantity() + "\t\t" + "$" + i.getPrice()*i.getQuantity());
			total=total+i.getPrice()*i.getQuantity();
		}
		System.out.println("Totals $ Amount in Cart: $"+total);
	}
	

	

}

