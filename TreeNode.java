import java.util.ArrayList;
import java.util.Iterator;

public class TreeNode<T> implements Iterator<TreeNode<T>>, Iterable<TreeNode<T>> {

	private T data;
	private TreeNode<T> parent = null;
	private ArrayList<TreeNode<T>> children = null;
	
	public TreeNode(T data) {
		this.data = data;
		children = new ArrayList<TreeNode<T>>();
	}
	
	public void setData(T data) {
		this.data = data;
	}
	
	public T getData() {
		return data;
	}
	
	public TreeNode<T> getParent() {
		return parent;
	}
	
	public void setParent(TreeNode<T> parent) {
		this.parent = parent;
	}
	
	public ArrayList<TreeNode<T>> getChildren() {
		return children;
	}
	
	public TreeNode<T> add(T data) {
		TreeNode<T> child = new TreeNode<T>(data);
		child.setParent(this);
		children.add(child);
		return child;
	}
	
	public boolean hasNext() {
		return children.listIterator().hasNext();
	}
	
	public TreeNode<T> next() {
		return children.listIterator().next();
	}
	
	public void remove() {
		children.listIterator().remove();
	}

	@Override
	public Iterator<TreeNode<T>> iterator() {
		// TODO Auto-generated method stub
		return this;
	}
}
