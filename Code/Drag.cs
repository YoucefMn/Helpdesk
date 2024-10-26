public class Drag
{
    private bool dragging = false;
    private Point dragCursorPoint;
    private Point dragFormPoint;
    private Form form;

    public Drag(Form form)
    {
        this.form = form;
    }

    public void Panel_MouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            dragging = true;
            dragCursorPoint = Cursor.Position;
            dragFormPoint = form.Location;
        }
    }
    public void Panel_MouseMove(object? sender, MouseEventArgs e)
    {
        if (dragging)
        {
            Point newCursorPoint = Cursor.Position;
            int xOffset = newCursorPoint.X - dragCursorPoint.X;
            int yOffset = newCursorPoint.Y - dragCursorPoint.Y;

            // Move the form according to the drag
            form.Location = new Point(dragFormPoint.X + xOffset, dragFormPoint.Y + yOffset);
        }
    }

    // MouseUp event handler
    public void Panel_MouseUp(object? sender, MouseEventArgs e)
    {
        dragging = false;
    }

}
