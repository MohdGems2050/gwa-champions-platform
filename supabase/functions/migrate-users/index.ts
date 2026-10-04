import { serve } from "https://deno.land/std@0.168.0/http/server.ts"
import { createClient } from "https://esm.sh/@supabase/supabase-js@2"

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const { users, adminSecret } = await req.json()

    // 1. Authenticate the migration request
    // The adminSecret must match the one set in the Supabase Edge Function secrets
    const expectedSecret = Deno.env.get('ADMIN_MIGRATION_SECRET')
    if (!expectedSecret || adminSecret !== expectedSecret) {
      return new Response(JSON.stringify({ error: 'Unauthorized: Invalid admin secret' }), {
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
        status: 403,
      })
    }

    if (!Array.isArray(users)) {
      return new Response(JSON.stringify({ error: 'Invalid payload: users array required' }), {
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
        status: 400,
      })
    }

    // 2. Initialize Supabase Admin Client
    // This uses the service_role key to bypass RLS and interact with the Auth API directly
    const supabaseAdmin = createClient(
      Deno.env.get('SUPABASE_URL') ?? '',
      Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? '',
      {
        auth: {
          autoRefreshToken: false,
          persistSession: false,
        }
      }
    )

    const results = {
      success: 0,
      failed: 0,
      errors: [] as string[],
    }

    // We need to map old localStorage IDs to new Auth UUIDs to maintain teacher_id relationships
    const idMap = new Map<string, string>()

    // Helper to process a user
    const processUser = async (user: any) => {
      try {
        const email = `${user.username}@gwa-champions.ae`
        
        // Check if user already exists
        const { data: existingUsers } = await supabaseAdmin.auth.admin.listUsers()
        const existing = existingUsers.users.find(u => u.email === email)
        
        let uid: string

        if (existing) {
          uid = existing.id
          // Optional: Update password if needed
          await supabaseAdmin.auth.admin.updateUserById(uid, { password: user.password })
        } else {
          // Create the Auth User
          const { data: authData, error: authError } = await supabaseAdmin.auth.admin.createUser({
            email: email,
            password: user.password,
            email_confirm: true,
          })
          
          if (authError) throw authError
          if (!authData.user) throw new Error("Failed to create auth user")
          
          uid = authData.user.id
        }

        idMap.set(user.id, uid)

        // Resolve teacher ID (if student)
        let resolvedTeacherId = null
        if (user.role === 'student' && user.teacherId) {
          resolvedTeacherId = idMap.get(user.teacherId) || null
        }

        // Upsert into profiles table
        const { error: profileError } = await supabaseAdmin
          .from('profiles')
          .upsert({
            id: uid,
            role: user.role,
            full_name: user.fullName,
            grade: user.grade,
            section: user.section || null,
            teacher_id: resolvedTeacherId
          })

        if (profileError) throw profileError
        
        results.success++
      } catch (err: any) {
        results.failed++
        results.errors.push(`Failed to migrate ${user.username}: ${err.message}`)
      }
    }

    // Process teachers first, so their UUIDs are available in the map for students
    const teachers = users.filter(u => u.role === 'teacher')
    const students = users.filter(u => u.role === 'student')

    for (const teacher of teachers) {
      await processUser(teacher)
    }

    for (const student of students) {
      await processUser(student)
    }

    return new Response(JSON.stringify(results), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 200,
    })
  } catch (error: any) {
    return new Response(JSON.stringify({ error: error.message }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 500,
    })
  }
})
